VERSION 5.00
Object = "{898CF0BC-AF45-45AC-9831-A27FC475455D}#1.0#0"; "WinXPctl.ocx"
Begin VB.Form frmSimulator 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   " SMS Web Gateway Simulator  .."
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSimulator.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   10290
   StartUpPosition =   2  '螢幕中央
   Begin VB.ComboBox cboOpt2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmSimulator.frx":0442
      Left            =   2070
      List            =   "frmSimulator.frx":044C
      TabIndex        =   24
      Text            =   "訂購 PPV 認證"
      Top             =   645
      Width           =   1935
   End
   Begin WinXP_Engine.WinXPctl xpc 
      Left            =   3510
      Top             =   -420
      _ExtentX        =   3995
      _ExtentY        =   873
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit (&X)"
      Height          =   375
      Left            =   6420
      TabIndex        =   22
      Top             =   8895
      Width           =   1335
   End
   Begin VB.Frame fraOpt 
      BackColor       =   &H00E0E0E0&
      Caption         =   " 傳 送 命 令 "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8655
      Left            =   150
      TabIndex        =   23
      Top             =   150
      Width           =   9795
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 28.ESERVICE回維修工單"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   225
         Index           =   27
         Left            =   5550
         TabIndex        =   34
         Top             =   7650
         Width           =   2925
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 27.ESERVICE回停拆移機單"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   225
         Index           =   26
         Left            =   5550
         TabIndex        =   33
         Top             =   7380
         Width           =   2925
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 26.ESERVICE回停拆移機單"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   225
         Index           =   25
         Left            =   5550
         TabIndex        =   32
         Top             =   7110
         Width           =   2805
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 25.ESERVICE回裝機單"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   225
         Index           =   24
         Left            =   5550
         TabIndex        =   31
         Top             =   6840
         Width           =   2295
      End
      Begin VB.ComboBox cboOpt24 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmSimulator.frx":0471
         Left            =   3060
         List            =   "frmSimulator.frx":047B
         TabIndex        =   30
         Text            =   "設備種類=1.MAC"
         Top             =   7590
         Width           =   1935
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 24. ETHOME-CM相關資訊查詢"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   225
         Index           =   23
         Left            =   150
         TabIndex        =   29
         Top             =   7650
         Width           =   4965
      End
      Begin VB.ComboBox cboOpt22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmSimulator.frx":04A0
         Left            =   3060
         List            =   "frmSimulator.frx":04AA
         TabIndex        =   28
         Text            =   "查帳務資訊"
         Top             =   6870
         Width           =   1935
      End
      Begin VB.ComboBox cboOpt23 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmSimulator.frx":04C6
         Left            =   3060
         List            =   "frmSimulator.frx":04D0
         TabIndex        =   27
         Text            =   "查帳務資訊"
         Top             =   7230
         Width           =   1935
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 23. CMCP相關資訊查詢"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   22
         Left            =   150
         TabIndex        =   26
         Top             =   7324
         Width           =   5025
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 22. CATV&DSTB相關資訊查詢"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   21
         Left            =   150
         TabIndex        =   25
         Top             =   7002
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 21. 經營區尋找"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   20
         Left            =   150
         TabIndex        =   20
         Top             =   6680
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 20. 欠費客戶名單查詢"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   19
         Left            =   150
         TabIndex        =   19
         Top             =   6358
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 19. 傳回SMS會員資料 ( For web update .. )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   18
         Left            =   150
         TabIndex        =   18
         Top             =   6036
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 13. 計次付費節目取消(需經過SMS訂購密碼認證,並包括CA系統訂購授權作業及處理結果)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   12
         Left            =   135
         TabIndex        =   12
         Top             =   4104
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 14. 計次付費節目取消原因資料"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   13
         Left            =   135
         TabIndex        =   13
         Top             =   4426
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00C0C0C0&
         Caption         =   " 15. 有線電視客戶基本資料查詢"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   225
         Index           =   14
         Left            =   135
         TabIndex        =   14
         Top             =   4748
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00C0C0C0&
         Caption         =   " 16. 有線客戶基本資料維護(少數基本資料欄位)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   225
         Index           =   15
         Left            =   135
         TabIndex        =   15
         ToolTipText     =   " 查詢(ppv,ippv) [提供個人與管理端] "
         Top             =   5070
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 17. 客戶安裝STB設備資料查詢(已安裝STB型號,編號及位置)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   16
         Left            =   135
         TabIndex        =   16
         Top             =   5392
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 18. 讀取點數收費資料"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   17
         Left            =   135
         TabIndex        =   17
         Top             =   5714
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 7. 客戶有線電視收視費轉帳作業/計次付費預付金額轉帳作業(信用卡小額付款機制)--線上付款授權"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   6
         Left            =   135
         TabIndex        =   6
         Top             =   2172
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 8. 點數線上訂購--線上付款授權"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   7
         Left            =   135
         TabIndex        =   7
         Top             =   2494
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 9. 點數線上訂購--下載點數"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   8
         Left            =   135
         TabIndex        =   8
         Top             =   2816
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 10. 客戶變更STB四碼數字Parental PIN Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   9
         Left            =   135
         TabIndex        =   9
         ToolTipText     =   " 查詢(ppv,ippv) [提供個人與管理端] "
         Top             =   3138
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 11. 客戶 iPPV訂購PIN Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   10
         Left            =   135
         TabIndex        =   10
         Top             =   3460
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 12. 計次付費節目訂購(需經過SMS訂購密碼認證,並包括CA系統訂購授權作業及處理結果)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   11
         Left            =   135
         TabIndex        =   11
         Top             =   3782
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 6. 客戶有線電視收視費轉帳作業/計次付費預付金額轉帳作業(信用卡小額付款機制)--線上付款授權--收查詢"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   5
         Left            =   135
         TabIndex        =   5
         Top             =   1850
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 5. 客戶帳務資料查詢(日期區間內得歷次帳單及繳款記錄)(個人)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   4
         Left            =   135
         TabIndex        =   4
         Top             =   1528
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 4. 客戶訂購計次付費節目明細查詢(ppv,ippv) [提供個人與管理端] "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   3
         ToolTipText     =   " 查詢(ppv,ippv) [提供個人與管理端] "
         Top             =   1206
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00C0C0FF&
         Caption         =   " 3. 計次付費節目內容資料 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   884
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 2. 訂購 PPV 認證 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   1
         Top             =   562
         Width           =   9500
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 1. 有線電視客戶身份 SMS 帳號密碼認證申請 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   9500
      End
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Testing (&E)"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   21
      Top             =   8895
      Width           =   1335
   End
End
Attribute VB_Name = "frmSimulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboOpt2_GotFocus()
    opt(1).Value = True
End Sub

Private Sub cboOpt23_GotFocus()
    opt(22).Value = True
End Sub

Private Sub Form_Load()
    With xpc
        .TextControl = False
        .OptionControl = False
        .ColorScheme = System
        .InitSubClassing
    End With
    cboOpt2.ListIndex = 0
    cboOpt22.ListIndex = 0
    cboOpt23.ListIndex = 0
End Sub

Public Function StartUp(ByVal strPara As String) As String
    StartUp = CreateObject("CS_Web.Gateway").Executing(strPara)
End Function

Private Sub cmdTest_Click()
    
    '1,2,19
    
    Dim bytFuncID As Byte
    Dim strResult As String
    bytFuncID = Val(GetOpt) + 1
    strResult = CallByName(Me, "Func" & bytFuncID, VbMethod)
    Debug.Print strResult
    MsgBox strResult, vbInformation, "Message"
    
End Sub

Public Function Func1() As String ' 有線電視客戶身份SMS帳號密碼認證申請 ( 命令 1 )
'   1、公司別、客戶編號、STB序號
'   會員代號、會員名稱、身分證字號、出生年月日、聯絡電話、
'   手機號碼、性別、聯絡地址、Email、PPV後付訂購密碼、登入密碼

'   1、公司別、客戶編號、STB序號
'   新的參數 : 客服網站會員代號、會員名稱、身分證字號、出生日期-年、出生日期-月、出生日期-日、
'   聯絡電話(宅) 、聯絡電話(公)、行動電話、性別、通訊地址-縣市、通訊地址-鄉鎮市區、通訊地址-路段街巷弄巷號樓、
'   Email、客服網站登入帳號、PPV訂購密碼、客服網站會員密碼、密碼提示問題、密碼提示答案、學歷、
'   願意收到有線電視服務及頻道節目等相關訊息、婚姻狀況、同住人數、同住成員、職業別、個人年收入、
'   身份別、帳號狀態、申請日期時間、啟用日期、最後登入日期、最後異動日期、系統台代碼

'1、公司別、客戶編號(折行)
'C、用戶姓名、電話(折行)
'D、DSTB設備序號、智慧卡卡號(折行)
'I、申請人、身份證、生日(折行)
'……

'1、公司別、客戶編號、種類、認證ID (當種類=0時認證ID須有值)(折行)
'C、用戶姓名、電話(折行)
'D、DSTB設備序號、智慧卡卡號(折行)
'I、申請人、身份證、生日(折行)
'……

'    Func1 = StartUp("1,1,12,0,0010000006" & vbCrLf & _
                                "C,用戶姓名,電話(折行)" & vbCrLf & _
                                "D,DSTB設備序號,智慧卡卡號(折行)" & vbCrLf & _
                                "I,申請人,身份證,生日(折行)")

'    Func1 = StartUp("1,1,12,0,0010000006,200605080002,test,,2005,08,04,,,,女,台中市,NO,第90路86號,winnie@eland.com.tw,null,,test,test,test,,,,,,資訊/高科技,30萬以下,1,Y,20050804,20050804,20050804,20050804,16,,,," & vbCrLf & _
                                "C,客戶名稱,4254389" & vbCrLf & _
                                "D,38383388,22660066" & vbCrLf & _
                                "I,沒有人申請,V111111111,19750922")
                                
'    Func1 = StartUp("1,1,12,1,,200605080002,test,A123456789,2005,08,04,,,,女,台中市,NO,第90路86號,winnie@123,null,888,test,test,test,,,,,,資訊/高科技,30萬以下,1,Y,20060509,20060509,20060509,20060509,16,,,," & vbCrLf & _
                                "D,119119374855,116021893636")
                                
'    Func1 = StartUp("1,1,12,1,,200605080002,金喜善,,2006,05,17,,,,女,,NO,A市第90路86號,winnie@eland.com.tw,,winnie,winnie,winnie,winnie,,,,,,資訊/高科技,30萬以下,1,Y,20060517,20060517,20060517,20060517,C,發票抬頭,發票種類,發票地址,統一編號" & vbCrLf & _
                                "D,119120020205,116025519411" & vbCrLf & _
                                "C,金喜善,987654321")
                                
'    Func1 = StartUp("1,16,1980,1,,200605080003,GAGA,A123456789,2005,08,04,,,,女,台中市,NO,第90路86號,winnie@eland.com.tw,null,888,test,test,test,,,,,,資訊/高科技,30萬以下,1,Y,20060509,20060509,20060509,20060509,16,,,," & vbCrLf & _
                                "C,客戶名稱,1980" & vbCrLf & _
                                "D,200411250066126,119119589500,116019114162")

    Func1 = StartUp("1,1,12,0,0010000006," & _
                                "MemberID,Hammer,V111111111," & _
                                "1975,09,22,05-3838338,07-5561926,0938123456,男," & _
                                "高雄市,左營區,博愛二路198號11樓之1," & _
                                "PowerHammer@msn.com,LoginID," & _
                                "PPVOrderPwd,LoginPassword,HintName,HintAnswer," & _
                                "幼稚園,1,未婚,200,一堆人,自由業,200萬,特殊身份,壯太," & _
                                "2005/05/01,2005/05/06,2005/05/20,2005/05/20,1," & _
                                "發票抬頭,發票種類,發票地址,統一編號,'000CE5DDF72C'" & _
                                "C,客戶名稱,4254389" & vbCrLf & _
                                "D,38383388,22660066" & vbCrLf & _
                                "I,沒有人申請,V111111111,19750922,3388")

                                
'    Func1 = StartUp("1,1,12,000CE5DDF72C," & _
                            "MemberID,Hammer,V111111111," & _
                            "1975,09,22,05-3838338,07-5561926,0938123456,男," & _
                            "高雄市,左營區,博愛二路198號11樓之1," & _
                            "PowerHammer@msn.com,LoginID," & _
                            "PPVOrderPwd,LoginPassword,HintName,HintAnswer," & _
                            "幼稚園,1,未婚,200,一堆人,自由業,200萬,特殊身份,壯太," & _
                            "2005/05/01,2005/05/06,2005/05/20,2005/05/20,1")

'    Func1 = StartUp("1,16,84,119120026164," & _
                                "0000001,會員一,A123456789," & _
                                "2005,07,20,02-1234556,02-1235679,0931234890,M," & _
                                "A市,,第一路182巷4弄2號3樓," & _
                                "abcd@test.com.tw,callcenter," & _
                                "ppvpwd,custpwd,密碼問題,密碼提示答案," & _
                                ",,,,,,,1,Y,20050720,20050720,20050720,20050720,16")
    
'    Func1 = StartUp("1,1,12,000CE5DDF72C," & _
                            "HaHa66,Hammer,V000000000," & _
                            "1975/09/22,07-5561926,0938123456,1," & _
                            "高雄市左營區博愛二路198號11樓之1," & _
                            "PowerHammer@msn.com,3388,2266")
    
'    Func1 = CreateObject("CS_MT.Birdge").StartUp("1,1,12,000CE5DDF72C," & _
'                            "HaHa66,Hammer,V000000000," & _
'                            "1975/09/22,07-5561926,0938123456,1," & _
'                            "高雄市左營區博愛二路198號11樓之1," & _
'                            "PowerHammer@msn.com,3388,2266")
    
'    Func1 = CreateObject("CS_Web.GateWay").Executing("1,1,12,000CE5DDF72C," & _
                            "HaHa66,Hammer,V000000000," & _
                            "1975/09/22,07-5561926,0938123456,1," & _
                            "高雄市左營區博愛二路198號11樓之1," & _
                            "PowerHammer@msn.com,3388,2266")
End Function

Public Function Func2() As String ' 訂購PPV認證 ( 命令 2 )
'   2、會員代號、認證ID、PPV後付訂購密碼

    '   舊參數 : 會員名稱、身分證字號、出生年月日、聯絡電話、手機號碼、性別、聯絡地址、Email、PPV後付訂購密碼、登入密碼
    '   新參數 : 客服網站會員代號、會員名稱、身分證字號、出生日期-年、出生日期-月、出生日期-日、
    '   聯絡電話(宅) 、聯絡電話(公)、行動電話、性別、通訊地址-縣市、通訊地址-鄉鎮市區、通訊地址-路段街巷弄巷號樓、
    '   Email、客服網站登入帳號、PPV訂購密碼、客服網站會員密碼、密碼提示問題、密碼提示答案、
    '   學歷、願意收到有線電視服務及頻道節目等相關訊息、婚姻狀況、同住人數、同住成員、職業別、
    '   個人年收入、身份別、帳號狀態、申請日期時間、啟用日期、最後登入日期、最後異動日期、
    '   系統台代碼(Web有會員之料被異動時需傳)
    '   發票抬頭、發票種類(二聯或三聯)、發票地址、統一編號(當發票種類為三聯時須為必填) 、設備流水號
'2,200605080001,0160000036,aaa,200605080001,test,,2005,08,04,,,,女,台中市,NO,第90路86號,winnie@eland.com.tw,null,,test,test,test,,,,,,資訊/高科技,30萬以下,1,Y,20050804,20050804,20050804,20050804,16,,,,,

    If cboOpt2.ListIndex = 0 Then
        Func2 = StartUp("2,MemberID,0010000006,PPVOrderPwd")
    Else
        Func2 = StartUp("2,MemberID,0010000006,PPVOrderPwd," & _
                                    "MemberID,Hammer,V111111111," & _
                                    "1975,09,22,05-3838338,07-5561926,0938123456,男," & _
                                    "高雄市,左營區,博愛二路198號11樓之1," & _
                                    "PowerHammer@msn.com,LoginID," & _
                                    "PPVOrderPwd,LoginPassword,HintName,HintAnswer," & _
                                    "幼稚園,1,未婚,200,一堆人,自由業,200萬,特殊身份,壯太," & _
                                    "2005/05/01,2005/05/06,2005/05/20,2005/05/20,1," & _
                                    "發票抬頭,,發票地址,統一編號',000CE5DDF72C'")
    End If
End Function

Public Function Func3() As String
End Function

Public Function Func4() As String ' 客戶訂購計次付費節目明細查詢(ppv,ippv) [提供個人與管理端] ( 命令 4 )
'   4、會員代號、認證ID、訂購起日、訂購迄日
        Func4 = StartUp("4,HaHa66,0010000062,20050410,20050412")
'   0:  成功
'   -5 : 查無訂購資料
End Function

Public Function Func5() As String ' 客戶帳務資料查詢(日期區間內得歷次帳單及繳款記錄)(個人) ( 命令 5 )
'   5、會員代號、認證ID、日期起日、日期迄日
        Func5 = StartUp("5,HaHa66,0010000022,20050101,20050412")
'   0:  成功
'   -6 : 查無帳務資料
End Function

Public Function Func6() As String ' 客戶有線電視收視費轉帳作業/計次付費預付金額轉帳作業(信用卡小額付款機制)--線上付款授權--收查詢 ( 命令 6 )
'   6、會員代號、認證ID
        Func6 = StartUp("6,HaHa66,0010000022")
'   0:  成功
'   -6 : 查無帳務資料
End Function

Public Function Func7() As String ' 客戶有線電視收視費轉帳作業/計次付費預付金額轉帳作業(信用卡小額付款機制)--線上付款授權 ( 命令 7 )
'   7、會員代號、認證ID、卡號、信用卡有效期限、信用卡背面授權碼、信用卡別(折行)
'   單據編號、項次(折行)
'   單據編號、項次(折行)
        Func7 = StartUp("7,HaHa66,0010000002,1111111111111111,122007,078038,1" & vbCrLf & _
                                    "200509ID0506570,1" & vbCrLf & _
                                    "200509ID0506570,2")
'        Func7 = StartUp("7,HaHa66,0010000002,1111111111111111,122007,,1")
'   0:  成功
'   -7 : 信用授權有誤
End Function

Public Function Func8() As String ' 點數線上訂購--線上付款授權 ( 命令 8 )
'   8、會員代號、認證ID、卡號、信用卡有效期限、信用卡背面授權碼、信用卡別(折行)
'   收費項目代號 (折行)
'   收費項目代號 (折行)
        Func8 = StartUp("8,HaHa66,0010000002,1111111111111111,122007,078038,1" & vbCrLf & _
                                    "235" & vbCrLf & _
                                    "301")
'        Func8 = StartUp("8,HaHa66,0010000002,1111111111111111,122007,,1")
'   0:  成功
'   -7 : 信用授權有誤
End Function

Public Function Func9() As String ' 點數線上訂購--下載點數 ( 命令 9 )
'    9、會員代號、認證ID(折行)
'    設備流水號、STB序號、ICC序號、點數(折行)
'    設備流水號、STB序號、ICC序號、點數(折行)
'    …
'        Func9 = StartUp("9,HaHa66,0010000022" & vbCrLf & "200406140057792,000CE5DDF72C,ICCsno,0")
        Func9 = StartUp("9,HaHa66,0010000022" & vbCrLf & "200406140057792,000CE5DDF72C,ICCsno,0" & vbCrLf)
'        Func9 = StartUp("9,HaHa66,0010000022" & vbCrLf & "200406140057792,000CE5DDF72C,ICCsno,0" & vbCrLf)
'        Func9 = StartUp("9,1126250439693,0160000028" & vbCrLf & "200502180068625,119121267688,116019066076,999")

'   0:  成功
'   -8 : 點數下載有誤
End Function

Public Function Func10() As String ' 客戶變更STB四碼數字Parental PIN Code ( 命令 10 )
'    10、會員代號、認證ID、設備流水號、STB序號、ICC序號
        Func10 = StartUp("10,HaHa66,0010000022,200406140057792,000CE5DDF72C,ICCsno")
'   0:  成功
'   -9 : 密碼變更失敗
End Function

Public Function Func11() As String ' 客戶 iPPV訂購PIN Code ( 命令 11 )
'   11、會員代號、認證ID、設備流水號、STB序號、ICC序號、iPPV PIN Code(4碼)
        Func11 = StartUp("11,HaHa66,0010000022,200406140057792,000CE5DDF72C,ICCsno,ABCD")
'   0:  成功
'   -9 : 密碼變更失敗
End Function

'cable_liga: API-12規格異動： 1．增加[價格]參數 2．小計金額、小計點數以意藍所傳參過來的回存。
'cable_liga: 5． 小計金額=節目價格+手續費 [意藍傳過來的價格]、 小計點數=EPG上之點數+點數 [意藍傳過來的點數]
'cable_liga: 設備流水號、STB序號、ICC序號、產品代碼、節目名稱、點數、價格(折行)

Public Function Func12() As String ' 計次付費節目訂購(需經過SMS訂購密碼認證,並包括CA系統訂購授權作業及處理結果) ( 命令 12 )
'   12、會員代號、認證ID、PPV後付訂購密碼、2(表來源網站) (折行)
'   設備流水號、STB序號、ICC序號、產品代碼、節目名稱、價格(折行)
'   設備流水號、STB序號、ICC序號、產品代碼、節目名稱、價格(折行)
'   ……..

'   節目訂購明細的第6個參數，由點數改傳為價格。參數個數不變
        Func12 = StartUp("12,HaHa66,0010000002,PPVOrderPwd,2" & vbCrLf & _
                                        "200509160058579,00111AC86A30,ICCsno,1,HaHa,200,2006/06/20 12:12:12" & vbCrLf & _
                                        "200512270058636,0002CF011E10,ICCsno,1,HaHa,400,2006/06/20 12:12:12")
'   0:  成功
'   -2:  訂購PPV認證錯誤
'   -4 : 未啟用PPV訂購權
'   -11 : 計次付費節目訂購失敗
End Function

Public Function Func13() As String ' 計次付費節目取消(需經過SMS訂購密碼認證,並包括CA系統訂購授權作業及處理結果) ( 命令 13 )
'   13、會員代號、認證ID、PPV後付訂購密碼、2(表來源網站)、取消時間(YYYYMMDD HH24MI)、取消人員代號、網站取消(取消人員)、訂單取消原因代碼、訂單取消原因(折行)
'   訂單單號、訂單項次(折行)
'   訂單單號、訂單項次(折行)
'   …
'        Func13 = StartUp("13,HaHa66,0010000002,PPVOrderPwd,2,20060119 2325,WEB,網站取消,38,不爽取消" & _
                                        "200000071,1" & vbCrLf & _
                                        "200000071,2")

        Func13 = StartUp("13,HaHa66,0010000002,PPVOrderPwd,2,20060119 2325,WEB,網站取消,38,不爽取消" & vbCrLf & _
                                        "2,1" & vbCrLf & _
                                        "3,2")

'   PS…取消人員代號='WEB'  ,   取消人員='網站取消'

'   0:  成功
'   -4 : 未啟用PPV訂購權
'   -12 : 計次付費節目取消訂購失敗

End Function

Public Function Func14() As String ' 計次付費節目取消原因資料 ( 命令 14 )
'   14
        Func14 = StartUp("14,HaHa66,0010000022")
'   0:  成功
'   -13 : 取消原因代碼讀取失敗
End Function

Public Function Func17() As String ' 客戶安裝STB設備資料查詢(已安裝STB型號,編號及位置) ( 命令 17 )
'   17、會員代號、認證ID
        Func17 = StartUp("17,HaHa66,0010000012")
'   0 : 成功
'   -14 : 查無安裝STB設備資料
End Function

Public Function Func18() As String ' 讀取點數收費資料 ( 命令 18 )
    '18、會員代號、認證ID
        Func18 = StartUp("18,HaHa66,0010000022")
'   0:  成功
'   -15 : 查無點數收費資料
End Function

Public Function Func19() As String ' 傳回SMS會員資料 For web update ( 命令 19 )
'   19、會員代號、認證ID
        Func19 = StartUp("19,HaHa66,0010000062")
'   0:  成功
'   -16 : 查無傳回SMS會員資料
End Function

Public Function Func20() As String ' 欠費客戶名單查詢 ( 命令 20 )
'   20、會員代號、認證ID
        Func20 = StartUp("20,HaHa66,0010000022")
'   0:  成功
'   -17 : 查無欠費資料
End Function

Public Function Func21() As String ' 經營區尋找 ( 命令 21 )
'   21、(系統台別複選)、縣市、鄉鎮市、村里庄埔路街其他、鄰、段、巷、弄、衖、號、其他
'        Func21 = StartUp("21,1,,中壢市,吉林路,1鄰,2巷,居易1弄,,2號,6樓之1")
'        Func21 = StartUp("21,1,您老師縣,中壢市,吉林路,1鄰,2巷,居易1弄,,2號,6樓之1")
'        Func21 = StartUp("21,1,,桃園市,大興西路,,1段,,,70號,")
        ' 桃園市大興西路1段70號
        Func21 = StartUp("21,2" & vbTab & "1,,中壢市,吉林路,1鄰,2巷,居易1弄,,2號,6樓之1")
'        Func21 = StartUp("21,1" & vbTab & "1,您老師縣,中壢市,吉林路,1鄰,2巷,居易1弄,,2號,6樓之1")
'   0:  成功
'   -18 : 查無經營區資料! 可能不在東森經營區或目前未配線至該地址
End Function

Public Function Func22() As String ' 客戶訂購數位頻道查詢 ( 命令 22 )
'   舊 :22, 公司別,申請人, 身份証字號, 電話, 種類(1=查帳務資訊、2=查設備資訊),起始日期,截止日期]
'   22、會員代號、認證ID、種類(1=查帳務資訊、2=查設備資訊)、起始日期、截止日期
    If cboOpt22.ListIndex <= 0 Then
        Func22 = StartUp("22,HaHa66,0050000002,1,20000101,20060922")
    Else
        Func22 = StartUp("22,HaHa66,0010000012,2,20000101,20060922")
    End If
'   0:  成功
'   -6 : 查無帳務資料
'   -14, 查無安裝STB設備資料
End Function

Public Function Func23() As String ' CMCP 相關資訊查詢 ( 命令 23 )
'   23、公司別、申請人、身份証字號、生日、種類(1=查帳務資訊、2=查設備資訊)、起始日期、截止日期
    If cboOpt23.ListIndex <= 0 Then
        Func23 = StartUp("23,1,廖運毅,A100000001,1977/01/03,1,2005/09/22,2006/09/22")
    Else
        Func23 = StartUp("23,1,王小明,V111111111,1978/09/03,2")
'        Func23 = StartUp("23,1,王小明,V111111111,1978/09/03,2,2005/09/22,2006/09/22")
    End If
'        Func23 = StartUp("23,1,陳清道,84177282,1949/08/18,1,2005/09/22,2006/09/22")
'        Func23 = StartUp("23,1,Hammer,V111111111,1975/09/22,2,2005/09/22,2006/09/22")
'   0:  成功
'   -6 : 查無帳務資料
'   -21, 查無安裝CMCP設備資料
End Function

Public Function Func24() As String ' ETHOME-CM相關資訊查詢 ( 命令 24 )
'   24、申請人姓名、身分種類、統編、設備種類、CM MAC

'   身分種類 ( 1:身分證;  2:統一編號)
'   身份證字號 或 統一編號,
'   設備種類 (1: MAC; 2: 帳號)
'   CM MAC 或 ACS 帳號

    If cboOpt24.ListIndex <= 0 Then
        ' 設備種類 MAC
        Func24 = StartUp("24,沈德宏,1,03-2530903,2,1111111111111111")
    Else
        ' 設備種類 帳號
        Func24 = StartUp("23,1,王小明,V111111111,1978/09/03,2")
'        Func24 = StartUp("23,1,王小明,V111111111,1978/09/03,2,2005/09/22,2006/09/22")
    End If
'   0:  成功
'   -99 : 參數不足
'   -20, 查無符合條件之資料
End Function


Public Function Func25() As String
'完工的部份
    Func25 = StartUp("25,5,9,200708II0032900,20070815 153050,,20070815 153050,0363,0363,20070815,")
'退單的部份
    'Func25 = StartUp("25,5,17,200708II0032899,,401,20070815 153050,0363,0363,20070815,")
End Function


Public Function Func26() As String
'完工的部份
    Func26 = StartUp("26,5,2410,200708PI0012041,20070815 153050,,20070815 153050,0363,0363,20070815,0,363")
'退單的部份
     'Func26 = StartUp("26,5,17,200708PI0012035,,101,20070815 153050,0363,0363,20070815,0,365")
End Function


Public Function Func28() As String
'完工的部份
    Func28 = StartUp("28,5,2410,200708PI0012041,20070815 153050,,20070815 153050,0363,0363,20070815,0,363")
'退單的部份
     'Func28 = StartUp("28,5,17,200708PI0012035,,101,20070815 153050,0363,0363,20070815,0,365")
End Function


'客戶 iPPV訂購PIN Code
Private Function GetOpt() As Byte
    Dim ctl As Control
    For Each ctl In Me
        If TypeOf ctl Is OptionButton Then
            If ctl.Value Then
                GetOpt = ctl.Index
                Exit For
            End If
        End If
    Next
    On Error Resume Next
    Set ctl = Nothing
End Function

Private Sub cmdExit_Click()
  On Error Resume Next
    Unload frmSimulator
    End
End Sub
Public Function Func27() As String ' 客戶有線電視收視費轉帳作業/計次付費預付金額轉帳作業(信用卡小額付款機制)--線上付款授權 ( 命令 7 )
'   7、會員代號、認證ID、卡號、信用卡有效期限、信用卡背面授權碼、信用卡別(折行)
'   單據編號、項次(折行)
'   單據編號、項次(折行)
        Func27 = StartUp("27,5,2838,06090470900")
'        Func7 = StartUp("7,HaHa66,0010000002,1111111111111111,122007,,1")
'   0:  成功
'   -7 : 信用授權有誤
End Function

