VERSION 5.00
Object = "{818C3FA5-C15E-4B09-B383-B06546ED97C7}#1.0#0"; "PHbtn.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   " 系 統 登 入"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txtPwd 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  '暫止
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   " 請輸入密碼 ! "
      Top             =   1080
      Width           =   2715
   End
   Begin PHbtn.PHbutton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2070
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "確 定 (&Y)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      FCOL            =   8388608
   End
   Begin PHbtn.PHbutton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   2070
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "取 消 (&N)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      FCOL            =   8388608
   End
   Begin VB.Label lblPwd 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "密 碼 :"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   630
      TabIndex        =   4
      Top             =   1185
      Width           =   495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "請 輸 入 密 碼 登 錄"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   780
      TabIndex        =   3
      Top             =   210
      Width           =   3015
   End
   Begin VB.Shape Shp 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '不透明
      BorderColor     =   &H00808080&
      FillColor       =   &H00400000&
      Height          =   1035
      Left            =   360
      Shape           =   4  '圓角矩形
      Top             =   765
      Width           =   3855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PR : Hammer
' Start date : 2004/12/23
' Last Modify : 2004/12/23

Option Explicit

Private bytCount As Byte

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
    End
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    With txtPwd
        If Len(Trim(.Text)) = 0 Then
            MsgBox "請輸入登入密碼 !", vbInformation, "訊息"
            On Error Resume Next
           .SetFocus
        Else
            If StrComp(.Text, strPassword) = 0 Then
                Hide
                frmMain.Show
                Unload Me
            Else
                If bytCount = 2 Then
                    MsgBox "請注意 , 您已登入三次錯誤 , 將無法使用本系統 !", vbInformation, "訊息"
                    Unload Me
                    End
                Else
                    bytCount = bytCount + 1
                    If StrComp(.Text, strPassword, vbTextCompare) = 0 Then
                        MsgBox "密碼錯誤 , 請重新輸入 ! ( 注意 : 密碼區分大小寫 ) ", vbInformation, "訊息"
                    Else
                        MsgBox "密碼錯誤 , 請重新輸入 !", vbInformation, "訊息"
                    End If
                    On Error Resume Next
                   .Text = ""
                   .SetFocus
                End If
            End If
        End If
    End With
  Exit Sub
ChkErr:
    ErrHandle Name, "Click Event : cmdOK_Click"
End Sub

Private Sub Form_Load()
    bytCount = 0
End Sub
