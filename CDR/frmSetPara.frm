VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSetPara 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "參數設定"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdEdit 
      Caption         =   "編輯"
      Height          =   465
      Left            =   1380
      TabIndex        =   14
      Top             =   4650
      Width           =   1065
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "儲存"
      Height          =   435
      Left            =   90
      TabIndex        =   13
      Top             =   4680
      Width           =   1185
   End
   Begin prjGiGridR.GiGridR ggrView 
      Height          =   2325
      Left            =   90
      TabIndex        =   12
      Top             =   2280
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   4101
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraConnData 
      Caption         =   "連線資訊"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   90
      TabIndex        =   5
      Top             =   1140
      Width           =   6585
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  '暫止
         Left            =   4560
         PasswordChar    =   "*"
         TabIndex        =   11
         Text            =   "EMCNCC"
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox txtUserId 
         Height          =   405
         Left            =   2670
         TabIndex        =   9
         Text            =   "EMCNCC"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtDataBase 
         Height          =   405
         Left            =   900
         TabIndex        =   7
         Text            =   "RDKNET"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblPass 
         Caption         =   "密碼"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblUser 
         Caption         =   "帳號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2220
         TabIndex        =   8
         Top             =   450
         Width           =   435
      End
      Begin VB.Label lblDataBase 
         Caption         =   "資料庫"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   450
         Width           =   675
      End
   End
   Begin VB.Frame fraAuto 
      Caption         =   "自動化參數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   6585
      Begin VB.TextBox txtAutoConnect 
         Alignment       =   1  '靠右對齊
         Height          =   345
         Left            =   4260
         TabIndex        =   3
         Text            =   "10"
         Top             =   390
         Width           =   585
      End
      Begin VB.TextBox txtAutoRun 
         Alignment       =   1  '靠右對齊
         Height          =   345
         Left            =   360
         TabIndex        =   1
         Text            =   "1"
         Top             =   390
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "秒斷線後自動重連"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5010
         TabIndex        =   4
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "分鐘後自動執行"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1110
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSetPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strConnectFile As String
Private strLine As String
Private Sub cmdSave_Click()
    rsReadFile.AddNew
    rsReadFile("AUTORUN") = txtAutoRun.Text
    rsReadFile("AUTOCONN") = txtAutoConnect.Text
    rsReadFile("DATABASE") = txtDataBase.Text
    rsReadFile("LOGINUSER") = txtUserId.Text
    rsReadFile("LOGINPASS") = txtPass.Text
    rsReadFile.Update
End Sub

Private Sub Form_Load()
    If Right(App.Path, 1) = "\" Then
        strConnectFile = App.Path & strConnectFileName
    Else
        strConnectFile = App.Path & "\" & strConnectFileName
    End If
    Call SetRsField
    If scr.FileExists(strConnectFile) Then
        Set FileStream = scr.OpenTextFile(strConnectFile, ForReading, False)
        'Call ReadParaData(FileStream)
        FileStream.Close
    End If
    Call subGrd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Set FileStream = scr.CreateTextFile(strConnectFile, True)
    Call WriteParaData(FileStream)
    rsReadFile.Close
    Set rsReadFile = Nothing
    FileStream.Close
    Set FileStream = Nothing
    Set scr = Nothing
End Sub


Private Sub WriteParaData(ByRef fs As TextStream)
    Dim i As Long
    Dim Xencrypt As Variant
    If rsReadFile.RecordCount > 0 Then
        rsReadFile.MoveFirst
    End If
    Do While Not rsReadFile.EOF
        strLine = Empty
        For i = 0 To rsReadFile.Fields.Count - 1
            strLine = strLine & rsReadFile(i).Value & IIf(i = rsReadFile.Fields.Count - 1, Empty, Chr(0))
        Next
        Xencrypt = CreateObject("Encryption.Password").Encrypt(strLine, "CS")
        fs.WriteLine Xencrypt
        rsReadFile.MoveNext
    Loop
    fs.Close
    Set fs = Nothing
End Sub
Private Sub subGrd()
    Dim mflds As New prjGiGridR.GiGridFlds
    With mflds
        .Add "DATABASE", , , , False, "資料庫", vbLeftJustify
        .Add "LOGINUSER", , , , False, "登入帳號", vbLeftJustify
    End With
    
    ggrView.AllFields = mflds
    ggrView.SetHead
    Set ggrView.Recordset = rsReadFile
    ggrView.Refresh
End Sub

