VERSION 5.00
Begin VB.Form frmTestSTB 
   BorderStyle     =   1  '單線固定
   Caption         =   "測試用"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   11.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "test.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4740
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdBatchTest 
      Caption         =   "CM 批次測試"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2910
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   1920
      Width           =   1605
   End
   Begin VB.CommandButton cmdPressureTest 
      Caption         =   "壓力測試"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   2400
      Width           =   1635
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2910
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   2400
      Width           =   1605
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2910
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   1260
      Width           =   1605
   End
   Begin VB.TextBox txtDataSource 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1470
      TabIndex        =   2
      Text            =   "rdknet"
      Top             =   840
      Width           =   3075
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  '暫止
      Left            =   1470
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "emcncc"
      Top             =   450
      Width           =   3075
   End
   Begin VB.TextBox txtUserId 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1470
      TabIndex        =   0
      Text            =   "emcncc"
      Top             =   60
      Width           =   3075
   End
   Begin VB.CommandButton cmdTEST 
      Caption         =   "CM 測試"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   1920
      Width           =   1635
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4567
      Y1              =   1760
      Y2              =   1760
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4567
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "DataSource"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   180
      TabIndex        =   7
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   180
      TabIndex        =   6
      Top             =   510
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "UserId"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   180
      TabIndex        =   5
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "frmTestSTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const lngCompCode = 5
Dim obj As Object
Dim garyGi(20) As String '--------- Dim Array --------------
Dim cnn As New ADODB.Connection

Private Sub cmdBatchTest_Click()
    Set obj = CreateObject("csControlPanelWeb.clsCMControl")
    Set obj.uconn = cnn
    obj.uErrPath = "C:\CSMISV40\ERRLOG"
    obj.uCompCodeFilter = lngCompCode
    obj.udefaultComp = lngCompCode
    obj.ucompcode = lngCompCode
    obj.ugarygi = JoinStr
    Set obj.uparentform = Me
    obj.SO1624AShow vbModal
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdICC_Click()
    obj.SO1622AShow
End Sub

Private Sub cmdPressureTest_Click()
    'obj.PressureTestShow
End Sub

Private Sub cmdSTB_Click()
    obj.SO1621AShow
End Sub

Private Sub cmdTEST_Click()
        Set obj = CreateObject("csControlPanelWeb.clsCMControl")
        Set obj.uconn = cnn
        obj.uErrPath = "C:\CSMISV40\ERRLOG"
        obj.uCompCodeFilter = lngCompCode
        obj.udefaultComp = lngCompCode
        obj.ucompcode = lngCompCode
        obj.uServiceType = "C"
        obj.ugarygi = JoinStr
        obj.uCustID = 166104
        Set obj.uparentform = Me
        obj.SO1623AShow
End Sub

'Select A.RowId, A.*
' FROM EMCNCC.So004 A
' Where A.CustId < 666666
' AND A.CompCode = 5
' AND A.FaciSNo is not null
' AND FaciCode In (Select CodeNo
' From EMCNCC.CD022
' WHERE RefNo in (2,5))
' AND A.PRDate is null
' ORDER BY A.InstDate

Private Sub cmdConnect_Click()
    On Error GoTo ChkErr
        garyGi(3) = "Provider=MSDAORA.1;" & _
                    "Password=" & txtPassword & ";" & _
                    "User ID=" & txtUserId & ";" & _
                    "Data Source=" & txtDataSource & ";" & _
                    "Persist Security Info=True"
        
        garyGi(3) = "Provider=OraOLEDB.Oracle.1;" & _
                    "Password=" & txtPassword & ";" & _
                    "User ID=" & txtUserId & ";" & _
                    "Data Source=" & txtDataSource & ";" & _
                    "Persist Security Info=True"
                    
        If cnn.State = adStateOpen Then cnn.Close
        cnn.Open garyGi(3)
        garyGi(4) = 0
        garyGi(16) = txtUserId

        cmdTEST.Enabled = True
        cmdBatchTest.Enabled = True
        cmdPressureTest.Enabled = True
    Exit Sub
ChkErr:
    MsgBox "錯誤:" & Err.Number & Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdExit.Value = True
End Sub

Private Sub Form_Load()
    garyGi(0) = "A"
    garyGi(1) = "Admin"
    
    garyGi(4) = "0"
    garyGi(5) = 132
    garyGi(6) = "2"
    garyGi(9) = lngCompCode
    garyGi(10) = "c:\CSMISV40\BIN\GICMIS1.INI"
    garyGi(11) = "c:\CSMISV40\BIN\GICMIS2.INI"
    garyGi(14) = lngCompCode
    
    garyGi(15) = "1,2,3,5,6,7,9"
End Sub

Private Function JoinStr() As String
    JoinStr = Join(garyGi, Chr(9))

End Function
