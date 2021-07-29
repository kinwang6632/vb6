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
      Text            =   "ksmis"
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
      Text            =   "cctv"
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
      Text            =   "cctv"
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
'Private Const lngCompCode = 1
Private Const lngCompCode = 7
'Private Const lngCompCode = 7 ' Previous
'Private Const lngCompCode = 3
Dim obj As Object
Dim garyGi(20) As String '--------- Dim Array --------------
Dim cnn As New ADODB.Connection

'Public Function ExecCmd(ByRef Gcn As Object, ByRef rs04 As Object, CmdID As String, _
'                                        AryGi As String, ErrPath As String, Optional ResvTime As String = "", _
'                                        Optional Transaction As Boolean = True, Optional RetMsg As String = "", _
'                                        Optional FaciCode As String = "", Optional FaciName As String = "", _
'                                        Optional FaciSno As String = "", Optional SNO As String = "", _
'                                        Optional MediaBillNo As String = "", Optional CmdSeqNo As String = "", _
'                                        Optional BaudRate As String = "", Optional ProbStopDate As String = "") As Boolean

Private Sub cmdBatchTest_Click()
    Dim o As Object
    Dim aSeqNo As String
    Dim rs04 As New ADODB.Recordset
    With rs04
        .CursorLocation = adUseClient
        
'        .Open "SELECT A.ROWID, A.* FROM GICMIS3.SO004 A" & _
                    " WHERE  A.CUSTID = 35 AND A.COMPCODE = 3 AND A.FACISNO IS NOT NULL" & _
                    " AND FACICODE IN (" & _
                    " SELECT CODENO FROM GICMIS3.CD022 WHERE REFNO IN (2,5)" & _
                    " ) ORDER BY A.PRDATE DESC,A.INSTDATE", cnn, adOpenKeyset, adLockOptimistic
                    
'       Previous
'        .Open "SELECT A.ROWID, A.* FROM CCTV.SO004 A" & _
                    " WHERE  A.CUSTID = 489786 AND A.COMPCODE = 7 AND A.FACISNO IS NOT NULL" & _
                    " AND FACICODE IN (" & _
                    " SELECT CODENO FROM CCTV.CD022 WHERE REFNO IN (2,5)" & _
                    " ) ORDER BY A.PRDATE DESC,A.INSTDATE", cnn, adOpenKeyset, adLockOptimistic
        
        .Open "SELECT A.ROWID, A.* FROM SO004 A" & _
                    " WHERE  A.CUSTID = 19 AND A.COMPCODE = 6 AND A.FACISNO IS NOT NULL" & _
                    " AND FACICODE IN (" & _
                    " SELECT CODENO FROM CD022 WHERE REFNO IN (2,5)" & _
                    " ) ORDER BY A.PRDATE DESC,A.INSTDATE", cnn, adOpenKeyset, adLockOptimistic
        
    End With
    Dim RetMsg As String
    Set o = CreateObject("csSPMconsole.clsCommand")
    ' C1
'    If o.ExecCmd(cnn, rs04, "A9", JoinStr, "C:\CSMISV40\ERRLOG", "20080202121212", , RetMsg, , , , "Sno", "MediaBillNo", "00720080123000000165") Then
'    If o.ExecCmd(cnn, rs04, "TRIAL", JoinStr, "C:\CSMISV40\ERRLOG", "20080202121212", , RetMsg, , , , _
'                            "Sno", "MediaBillNo", "00720080123000000165", "11") Then
    If o.ExecCmd(cnn, rs04, "A1_PIP", JoinStr, "C:\CSMISV40\ERRLOG", "", , RetMsg, , , , _
                            "Sno", "MediaBillNo", aSeqNo, "12", "2008/02/02") Then
        MsgBox "OK!"
    Else
        MsgBox RetMsg
    End If
    MsgBox aSeqNo
'    Set obj = CreateObject("csControlPanelWeb.clsCMControl")
'    Set obj.uconn = cnn
'    obj.uErrPath = "C:\CSMISV40\ERRLOG"
'    obj.uCompCodeFilter = lngCompCode
'    obj.udefaultComp = lngCompCode
'    obj.ucompcode = lngCompCode
'    obj.ugarygi = JoinStr
'    Set obj.uparentform = Me
'    obj.SO1624AShow vbModal
    
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
    Dim rsNew As New ADODB.Recordset
    Dim aSQL As String
    aSQL = "SELECT * FROM SO004 WHERE SEQNO='200710230255272'"
    
    rsNew.CursorLocation = adUseClient
    rsNew.Open aSQL, cnn, adOpenKeyset, adLockOptimistic
    Set obj = CreateObject("csSPMconsole.clsCMControl")
    With obj
        Set .uconn = cnn
        .uErrPath = "C:\CSMISV40\ERRLOG"
        .uCompCodeFilter = lngCompCode
        .udefaultComp = lngCompCode
        .ucompcode = lngCompCode
        .uNewBaudRate = "4"
        .uServiceType = "I"
        .ugarygi = JoinStr
'        .uProcessType = "A4"
'        .uProcessType = "C1"
'        .uProcessType = "TRIAL"
'        .uProcessType = "STOPTRIAL"
'        .uProcessType = "A1_BU"
'        .uProcessType = "A1_PIP"
'        .uProcessType = "RP"
'        .uProcessType = "A7"
'        .uCustID = 489786 ' 401950 ' 35 ' 165
'        .uCustID = 275
'        .uDefaultRowID = "AABCuuAAoAAACJAAAE"
'        .uSno = "200803II0187163"
'        .uIPnum = 2
'        .uIP = "127.0.0.1"
'        .uCPE = "000B063D8C74"
'        .uMediaBillNo = "MediaBillNo"
        .uResvTime = "20171215121212"
'        .uBuadRate = "14"
'        .uBuadRate = "4"
'        .uProbStopDate = "2008/02/02"
        .uBuadRate = "2"
'        .uFaciCode = "206"
'        .uFaciSno = "3388"
'        .uCustID = 421 ' 401950 ' 35 ' 165
'        .uCustID = 35 ' 401950 ' 35 ' 165
'        .uCustID = 8
'        .uCustID = 165
'        .uCustID = 4
'        .uCustID = 401963 ' Previous
        .uCustID = 69
'        .uOldIP = "127.0.0.1"
'        .uIP = "127.0.0.3"
        .uOldCPE = "00:90:96:07:9A:4A"
        .uCPE = "00:D0:59:1F:46:89"
        .uNewIP = "168.95.1.1"
        Set .uUpdRs = rsNew
'       .uCustID=275
'        .uSendKey = True
        Set .uparentform = Me
        .SO1623AShow 1, Nothing
        MsgBox .uCmdSeqNo
    End With
    On Error Resume Next
    rsNew.Close
    Set rsNew = Nothing
End Sub

Private Sub cmdConnect_Click()
    On Error GoTo ChkErr
        garyGi(3) = "Provider=MSDAORA.1;" & _
                    "Password=" & txtPassword & ";" & _
                    "User ID=" & txtUserId & ";" & _
                    "Data Source=" & txtDataSource & ";" & _
                    "Persist Security Info=True"
        
'        garyGi(3) = "Provider=OraOLEDB.Oracle.1;" & _
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
    garyGi(6) = "3"
    garyGi(9) = lngCompCode
    garyGi(10) = "c:\CSMISV40\BIN\GICMIS1.INI"
    garyGi(11) = "c:\CSMISV40\BIN\GICMIS2.INI"
    garyGi(14) = lngCompCode
    
    txtDataSource = "RDKNET"
    txtUserId = "tbcsh"
    txtPassword = "tbcsh"
    
    garyGi(15) = "1,2,3,5,6,7,9"
    
'    txtUserId = "gicmis3"
'    txtPassword = "may"
'    txtDataSource = "pikachu"
    
'    txtUserId = "tbcnty"
'    txtPassword = "tbcnty"
'    txtDataSource = "eserver"
    'TBCCCTV/TBCCCTV
    cmdConnect.Value = True
    'cmdTEST.Value = True
End Sub

Private Function JoinStr() As String
    JoinStr = Join(garyGi, Chr(9))
End Function
