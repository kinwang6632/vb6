VERSION 5.00
Object = "{818C3FA5-C15E-4B09-B383-B06546ED97C7}#1.0#0"; "PHbtn.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  '�S���ؽu
   Caption         =   "�t�εn�J [SO0000A]"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Login.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Login.frx":0442
   ScaleHeight     =   3885
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PHbtn.PHbutton cmdOk 
      Default         =   -1  'True
      Height          =   345
      Left            =   2730
      TabIndex        =   4
      Top             =   2800
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "�T  �w (&Y)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12640511
      FCOL            =   12582912
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  '�Ȥ�
      Left            =   4260
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2220
      Width           =   2385
   End
   Begin VB.TextBox txtUserId 
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4260
      TabIndex        =   1
      Top             =   1830
      Width           =   2385
   End
   Begin PHbtn.PHbutton cmdCancel 
      Height          =   345
      Left            =   4710
      TabIndex        =   5
      Top             =   2800
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "��  �� (&N)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12640511
      FCOL            =   12582912
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Left            =   90
      Picture         =   "Login.frx":10E7F
      Stretch         =   -1  'True
      Top             =   30
      Width           =   420
   End
   Begin VB.Label lblTestShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�t�εn�J [SO0000A]"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   600
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image imgMinButton 
      Height          =   285
      Left            =   3660
      Picture         =   "Login.frx":11B49
      Stretch         =   -1  'True
      Top             =   4050
      Width           =   285
   End
   Begin VB.Image imgTop 
      Height          =   385
      Left            =   210
      Picture         =   "Login.frx":120CD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   465
   End
   Begin VB.Image imgDisabled 
      Height          =   315
      Index           =   0
      Left            =   1320
      Picture         =   "Login.frx":12B09
      Top             =   4020
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDisabled 
      Height          =   315
      Index           =   1
      Left            =   1680
      Picture         =   "Login.frx":1308D
      Top             =   4020
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDisabled 
      Height          =   315
      Index           =   2
      Left            =   2040
      Picture         =   "Login.frx":13611
      Top             =   4020
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgEnabled 
      Height          =   315
      Index           =   0
      Left            =   1320
      Picture         =   "Login.frx":13B95
      Top             =   4380
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgEnabled 
      Height          =   315
      Index           =   1
      Left            =   1680
      Picture         =   "Login.frx":14119
      Top             =   4380
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgEnabled 
      Height          =   315
      Index           =   2
      Left            =   2040
      Picture         =   "Login.frx":1469D
      Top             =   4380
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDimmed 
      Height          =   315
      Index           =   0
      Left            =   1320
      Picture         =   "Login.frx":14C21
      Top             =   4740
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDimmed 
      Height          =   315
      Index           =   1
      Left            =   1680
      Picture         =   "Login.frx":151A5
      Top             =   4740
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgDimmed 
      Height          =   315
      Index           =   2
      Left            =   2040
      Picture         =   "Login.frx":15729
      Top             =   4740
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgShutDown 
      Height          =   360
      Index           =   0
      Left            =   1320
      Picture         =   "Login.frx":15CAD
      Top             =   5100
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgShutDown 
      Height          =   360
      Index           =   1
      Left            =   1920
      Picture         =   "Login.frx":163B1
      Top             =   5100
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgLogOff 
      Height          =   360
      Index           =   0
      Left            =   1320
      Picture         =   "Login.frx":16AB5
      Top             =   5460
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgLogOff 
      Height          =   360
      Index           =   1
      Left            =   1920
      Picture         =   "Login.frx":171B9
      Top             =   5460
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgTopLeft 
      Height          =   385
      Left            =   0
      Picture         =   "Login.frx":178BD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblTest 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�t�εn�J [SO0000A]"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   180
      Left            =   600
      TabIndex        =   9
      Top             =   105
      Width           =   1575
   End
   Begin VB.Image imgLeft 
      Height          =   495
      Left            =   4590
      Picture         =   "Login.frx":17D89
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   45
   End
   Begin VB.Image imgTopRight 
      Height          =   390
      Left            =   6810
      Picture         =   "Login.frx":18025
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
   Begin VB.Image imgRight 
      Height          =   375
      Left            =   4710
      Picture         =   "Login.frx":184F1
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   45
   End
   Begin VB.Image imgBottom 
      Height          =   45
      Left            =   3630
      Picture         =   "Login.frx":1878D
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   630
   End
   Begin VB.Image imgExit 
      Height          =   285
      Left            =   3990
      Picture         =   "Login.frx":18951
      Stretch         =   -1  'True
      Top             =   4050
      Width           =   285
   End
   Begin VB.Image imgMaxButton 
      Height          =   285
      Left            =   3330
      Picture         =   "Login.frx":18ED5
      Stretch         =   -1  'True
      Top             =   4050
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblCompName 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3750
      TabIndex        =   8
      Top             =   990
      Width           =   3045
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   8.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   6255
      TabIndex        =   7
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�ϥΪ̱K�X(&P)"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2115
      TabIndex        =   2
      Top             =   2265
      Width           =   1950
   End
   Begin VB.Label lblUserId 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�ϥΪ̥N��(&I)"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2115
      TabIndex        =   0
      Top             =   1845
      Width           =   1950
   End
   Begin VB.Label lblSysName 
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2085
      TabIndex        =   6
      Top             =   1395
      Width           =   4005
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gstrStationCount As String
Dim intX As Integer
Dim OpenDB As New giOpenDBObj.OpenDBObj
Dim ServerStr As String
Dim PortStr As String

Dim mouseX, mouseY As Integer

Private Sub Form_Activate()
  On Error GoTo ChkErr
    With Me
        .Move (Screen.Width - .Width) / 2, ((Screen.Height - .Height) / 2)
    End With
    UpdateTime
    CsMisUpdate
'    Call ReadINI2
'    On Error Resume Next
'    txtUserId.SetFocus
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Avtivate"
End Sub

Private Sub UpdateTime()
  On Error GoTo ChkErr
    If Val(gcnGi.Execute("SELECT SYSTIME FROM " & GetOwner & "SO041 WHERE ROWNUM <=1 ORDER BY COMPCODE").GetString(adClipString, , , , 1) & "") = 1 Then
        Date = gcnGi.Execute("SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD') FROM DUAL").GetString & ""
        Time = Trim(gcnGi.Execute("SELECT TO_CHAR(SYSDATE,'HH24:MI:SS') FROM DUAL").GetString & "")
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "UpdateTime"
End Sub

Public Sub CsMisUpdate()
  On Error GoTo ChkErr
    Dim strVal As String
    Dim UpdTimer As Long
    On Error Resume Next
    If Alfa Then Exit Sub
    AppActivate "PrjNotify"
    If Err.Number = 0 Then Exit Sub
    On Error GoTo ChkErr
    UpdTimer = Val(gcnGi.Execute("SELECT CSMISUPDATE FROM " & GetOwner & "SO041 WHERE ROWNUM <=1 ORDER BY COMPCODE").GetString & "")
    If UpdTimer = 0 Then Exit Sub
    strVal = ReadFromINI2("SOEXEPATH") & "\CSUpDate.EXE"
    If Dir(strVal) = "" Then
        MsgBox strVal & " ���s�b!!", , "�ާ@���~!!"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Shell strVal & " " & UpdTimer, vbNormalFocus
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
    ErrSub Me.Name, "NormalExe"
End Sub

Private Sub Form_Initialize()
  On Error GoTo ChkErr
    gstrCurDir = CurPath
    Call ReadCDRINI  ' �� GICMIS2.INI Ū����T
    If OpenDB.OpenConnection(gcnGi, garyGi(3), True, False, True) <> giOK Then     ' �إ߻P��Ʈw���s�u
        If OpenDB.OpenConnection(gcnGi, garyGi(3), True, False, True) <> giOK Then
            If OpenDB.OpenConnection(gcnGi, garyGi(3), True, False, True) <> giOK Then
                If Err.Number <> 0 Then MsgBox Err.Description, vbInformation, Err.Number
                Select Case True
                       Case InStr(1, OpenDB.GetErrDesc, "01017")
                            MsgBox "��Ʈw [UID] �� [�K�X] ���~ !", vbInformation, "�T��"
                       Case InStr(1, OpenDB.GetErrDesc, "12154")
                            MsgBox "��Ʈw [DSN �A�ȦW��]  ���~ !", vbInformation, "�T��"
                       Case InStr(1, OpenDB.GetErrDesc, "12514")
                            MsgBox "Net 8 �A�ȦW�ٳ]�w���~ !", vbInformation, "�T��"
                       Case InStr(1, OpenDB.GetErrDesc, "12541")
                            MsgBox "Net 8 �D�q���W�ٳ]�w���~�ΨS����ť�� !", vbInformation, "�T��"
                       Case InStr(1, OpenDB.GetErrDesc, "12560")
                            MsgBox "�������q�κ����L�k�s�ܸ�Ʈw���A�� !", vbInformation, "�T��"
                       Case Else
                            MsgBox OpenDB.GetErrDesc, vbInformation, OpenDB.GetErrNum
                End Select
                End
            End If
        End If
    Else
        If gcnGi.State = 1 Then
            PrintErrFlag = gcnGi.Execute("SELECT ERRLOGRPT FROM " & GetOwner & "SO041").GetString(adClipString, 1) = 1
        Else
            PrintErrFlag = False
        End If
    End If
    Screen.MousePointer = vbDefault
    ChkMonitorPixel
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Initialize"
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next
    Dim hwnd As Long
    hwnd = FindWindow(vbNullString, "PrjNotify")
    If hwnd <> 0 Then
        SetForegroundWindow hwnd
        PostMessage hwnd, &H10, 0&, 0&
    End If
    Unload Me
    End
End Sub

Private Sub cmdOk_Click()
  On Error GoTo ChkErr
    If IsDataOk Then CheckPassword         ' �ˬd��J��ƬO�_���T
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdOk_Click"
End Sub

Private Function IsDataOk() As Boolean ' �ˬd���n���O�_����, �Y�����n���B�L��, �h���"��쥲�ݦ���", �Bfocus��������W
  On Error GoTo ER
    IsDataOk = False
    If MustExist(txtUserId, , "�ϥΪ̥N��") = False Then Exit Function
    If MustExist(txtPassword, , "�ϥΪ̱K�X") = False Then Exit Function
    IsDataOk = True
  Exit Function
ER:
    ErrSub Me.Name, "IsDataOk"
End Function

Public Function CheckPassword() ' �ˬd��J��T�O�_���T
  On Error GoTo ChkErr
    Dim rsChkPwd As New ADODB.Recordset
    Dim strSQL As String
    Screen.MousePointer = vbHourglass
        garyGi(15) = ""

    If gcnGi.State = adStateClosed Then OpenDB.OpenConnection gcnGi, garyGi(3) ' �إ߻P��Ʈw���s�u
    If txtUserId = "POWERUSER" Then
        Dim PUpwd As String
        On Error Resume Next
        PUpwd = gcnGi.Execute("SELECT IDPWD FROM " & GetOwner & "SOPOWER").GetString(adClipString, 1, "", "", "")
        If PUpwd = Empty Then
            MsgBox "��ƪ�� [SOPOWER] �]�w�����T !!", vbInformation, "�T��"
        Else
            If Xdecrypt(PUpwd, "CS") <> txtPassword Then
                garyGi(4) = -1
            Else
                garyGi(4) = "S"
                garyGi(15) = gcnGi.Execute("SELECT CODENO FROM " & GetOwner & "CD039 ORDER BY CODENO").GetString(adClipString, , "", ",")
                garyGi(15) = Left(garyGi(15), Len(garyGi(15)) - 1)
            End If
        End If
    Else
        strSQL = "SELECT * FROM " & GetOwner & "SO026 WHERE UserID ='" & txtUserId & "' AND " & _
                 "Password = " & "'" & Xencrypt(txtPassword, "CS") & "'"
        OpenRecordset rsChkPwd, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False
        If rsChkPwd.RecordCount = 0 Then ' �A�ˬd�� �ϥΪ� �O�_�� Supervisor
            strSQL = "SELECT COUNT(*) FROM " & GetOwner & "SO041 WHERE MTUID ='" & Xencrypt(txtUserId.Text, "CS") & "' AND " & _
                     "MTUPSWD ='" & Xencrypt(txtPassword.Text, "CS") & "'"
            OpenRecordset rsChkPwd, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False
            If rsChkPwd(0) > 0 Then
                garyGi(4) = 0
                garyGi(15) = gcnGi.Execute("SELECT CODENO FROM " & GetOwner & "CD039 ORDER BY CODENO").GetString(adClipString, , "", ",")
                garyGi(15) = Left(garyGi(15), Len(garyGi(15)) - 1)
            Else
                garyGi(4) = -1
            End If
        Else
            garyGi(4) = rsChkPwd("GroupID")
            garyGi(15) = rsChkPwd("COMPSTR")
        End If
    End If
    Screen.MousePointer = vbDefault
    If garyGi(4) = "S" Then GoTo 38
    If garyGi(4) < 0 Then  ' �p�G��J��Ƥ��ŦX
        intX = intX + 1
        Select Case intX
            Case 1
                MsgBox " �ϥΪ̥N���αK�X����,���ˬd�]�w!! ! ", vbInformation, " �K�X�n�J���~ !! "
                'MsgBox " �ϥΪ̥N���αK�X����,���ˬd�]�w!! ! (�z�٦� 2 �����|!!) ", vbInformation, " �K�X�n�J���~ !! (�Ĥ@��)"
                'Call WrongPwd
                Unload Me
                End
'            Case 2
'                MsgBox " �ϥΪ̥N���αK�X����,�н��ˬd�]�w!! (�z�ȳ� 1 �����|!!) ", vbInformation, " �K�X�n�J���~ !! (�ĤG��)"
'                'Call WrongPwd
'                Unload Me
'                End
'            Case 3
'                MsgBox " ��p ! �K�X�n�J 3 �����~ !! �z�L�v���ϥΥ��t�� !! BYE~BYE ! ", vbExclamation, " �t�εn�J���� , �Y�N���� ! "
'                Unload Me
'                End
        End Select
    Else ' ��J��ƲŦX
38:
        On Error GoTo ChkErr
        If IsNumeric(garyGi(4)) Then
            If garyGi(4) = 0 Then
                garyGi(0) = txtUserId.Text
                garyGi(1) = "Admin"
                'garyGi(2) = txtUserId.Text
            Else
                garyGi(0) = txtUserId
                garyGi(1) = rsChkPwd("UserName")
                'garyGi(2) = rsChkPwd("UserId")
            End If
        Else
            If garyGi(4) = "S" Then
                garyGi(0) = txtUserId.Text
                garyGi(1) = "MSO"
                'garyGi(2) = txtUserId.Text
            Else
                MsgBox "�L�k�T�{���� ! �L�v�ϥΥ��t�� !", vbInformation, "�T��"
                Unload Me
                End
                Exit Function
            End If
        End If
'        If garyGi(4) = "S" Then
'            garyGi(0) = txtUserId.Text
'            garyGi(1) = "MSO"
'            'garyGi(2) = txtUserId.Text
'        ElseIf garyGi(4) = 0 Then
'            garyGi(0) = txtUserId.Text
'            garyGi(1) = "Admin"
'            'garyGi(2) = txtUserId.Text
'        Else
'            garyGi(0) = txtUserId
'            garyGi(1) = rsChkPwd("UserName")
'            'garyGi(2) = rsChkPwd("UserId")
'        End If
'       garyGi(4) = rsChkPwd("GroupID")
'       If Alfa Then garyGi(4) = "0"
        If rsChkPwd.State = 1 Then rsChkPwd.Close
        Set rsChkPwd = Nothing
        If IsNumeric(garyGi(4)) Then
            If garyGi(4) = 0 And Not ChkComputerName(True) Then Exit Function
        Else
            If Not ChkComputerName(True) Then Exit Function
        End If
'        If garyGi(4) = 0 Or garyGi(4) = "S" Then
'            If Not ChkComputerName(True) Then Exit Function
'        Else
'            If Not ChkComputerName(False) Then Exit Function
'        End If
'       If Not ChkComputerName(IIf(garyGi(4) = 0, True, False)) Then Exit Function
        Crm = gcnGi.Execute("SELECT USECRM FROM " & GetOwner & "SO041 WHERE COMPCODE=" & gEntryComp).GetString = 1
        Crm = UCase(Command) = "CRM" And Crm
        If Crm Then
            If MonitorPixel <> "1024x768" Then MsgBox "�] CRM �{����������� 1024 x 768 ���ù� Pixel ..!! �нվ� ..", vbInformation, "�T��": Unload Me: End
        End If
        garyGi(12) = txtPassword.Text
        Me.Hide
        SetDataArea
'       txtUserId.Text = ""
        txtPassword.Text = ""
'       MsgBox gcnGi.Execute("SELECT MID FROM SO029 WHERE GROUP" & garyGi(4) & "=1 And COMPCODE=" & GetNullString(gCompCode, giLongV)).GetString
        GetUserPriv "", "", True
        
        frmMain.Show 1
        'frmFind.Show
    End If
  Exit Function
ChkErr:
    ErrSub Me.Name, "CheckPassword"
End Function

Private Sub SetDataArea()
  On Error GoTo ChkErr
    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = gcnGi
        .CommandText = "SF_GETDATAAREA"
        .Parameters.Append .CreateParameter("RetVal", adVarChar, adParamReturnValue, 2000)
        .Parameters.Append .CreateParameter("UID", adVarChar, adParamInput, 2000, garyGi(0))
'        .Parameters.Append .CreateParameter("CompCode", adVarNumeric, adParamInput, 2000, gCompCode)
        .Execute
         garyGi(2) = .Parameters("RetVal") & ""
    End With
'   garyGi(2) = CreateObject("csDMToolsSF.clsStoreFunction").SF_GETDATAAREA(gcnGi, garyGi(0))
  Exit Sub
ChkErr:
    If Err.Number = -2147217900 Then
        With cmd
            .Parameters.Append .CreateParameter("CompCode", adVarNumeric, adParamInput, 2000, gCompCode)
            .Execute
             garyGi(2) = .Parameters("RetVal") & ""
        End With
    Else
        ErrSub Me.Name, "SetDataArea"
    End If
End Sub

' ���o���x�q���W�١A�m�� gstrPcName
Private Function ChkComputerName(PowerUser As Boolean) As Boolean
  On Error GoTo ChkErr
    
    Dim rsChkStation As New ADODB.Recordset
    Dim rsOnLine As New ADODB.Recordset
    Dim CompCode As Integer
    Dim strSQL As String
    Dim CanGo As Boolean
    gstrPcName = CreateObject("WScript.Network").ComputerName
    
    strSQL = "SELECT * FROM " & GetOwner & "SO025 WHERE STATIONNAME = " & "'" & gstrPcName & "'"
    
    rsChkStation.CursorLocation = adUseClient
    OpenRecordset rsChkStation, strSQL, gcnGi, adOpenKeyset, adLockOptimistic, adUseClient, False
    
    If rsChkStation.EOF Then ' �p�G�W�����`���u
        rsOnLine.CursorLocation = adUseServer
        OpenRecordset rsOnLine, "SELECT COUNT(*) FROM " & GetOwner & "SO025", gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False
        
        If PowerUser Then ' �g�J�������n�J�O��
            gEntryComp = GetRsValue("SELECT COMPCODE FROM " & GetOwner & "SO041 ORDER BY SYSID", gcnGi, "�b�}�� [�t�η~�̰Ѽ���](SO041) ����Ʈɵo�Ϳ��~")
            garyGi(9) = gEntryComp
            CompCode = gEntryComp
            gCompCode = gEntryComp
        Else
            CompCode = gcnGi.Execute("SELECT COMPCODE FROM " & GetOwner & "SO026 WHERE USERID='" & txtUserId & "'").GetString
            gEntryComp = CompCode
            garyGi(9) = CompCode
            gCompCode = CompCode
        End If
        
        CanGo = ChkDupLogin
        If CanGo Then
            gcnGi.Execute "INSERT INTO " & GetOwner & "SO025 (STATIONNAME,USERID,LOGINTIME,COMPCODE) " & _
                          " VALUES ('" & gstrPcName _
                           & "','" & txtUserId & _
                           "',TO_DATE(" & _
                           Format(Now, "YYYYMMDDHHMMSS") & _
                           ",'YYYYMMDDHH24MISS')," & CompCode & ")"
            ChkComputerName = True
        Else
            ChkComputerName = CanGo
        End If
    
    Else ' �p�G�W�������`���u
        If PowerUser Then
            gEntryComp = GetRsValue("SELECT COMPCODE FROM " & GetOwner & "SO041 ORDER BY SYSID", gcnGi, "�b�}�� [�t�η~�̰Ѽ���](SO041) ����Ʈɵo�Ϳ��~")
            garyGi(9) = gEntryComp
            CompCode = gEntryComp
            gCompCode = gEntryComp
        Else
            CompCode = gcnGi.Execute("SELECT COMPCODE FROM " & GetOwner & "SO026 WHERE USERID='" & txtUserId & "'").GetString
            gEntryComp = CompCode
            garyGi(9) = CompCode
            gCompCode = CompCode
        End If
        ChkComputerName = ChkDupLogin
        If Not Alfa Then MsgBox "�`�N! �ϥΪ� " & rsChkStation("UserID") & " (" & garyGi(1) & ") �W�������`���u..!!", vbExclamation, "ĵ�i�T��.."
       ' ��W�@���n�J�O����s���������n�J�ɶ�
        gcnGi.Execute "UPDATE " & GetOwner & "SO025 SET LOGINTIME = TO_DATE(" & Format(Now, "YYYYMMDDHHMMSS") & ",'YYYYMMDDHH24MISS') WHERE StationName like '" & gstrPcName & "'"
    End If
    
    'UseCallCenter = gcnGi.Execute("SELECT CHECKCALL FROM SO041 WHERE COMPCODE=" & garyGi(9) & "").GetString & ""
'    If UseCallCenter Then
'        Dim vCCinfo As Variant
'        vCCinfo = Split(gcnGi.Execute("SELECT SERVERNAME,PORT FROM SO041 WHERE COMPCODE=" & garyGi(9) & "").GetString, vbTab)
'        ServerStr = Trim(vCCinfo(0) & "")
'        PortStr = Trim(vCCinfo(1) & "")
'        If Right(PortStr, 1) = Chr(13) Then PortStr = Left(PortStr, Len(PortStr) - 1)
'        If ServerStr = "" Or PortStr = "" Then
'            UseCallCenter = False
'            MsgBox "�Y�ϥ� Call Center �Ҳ�," & strCrLf & "�Щ�SO041�]�wServerName & Port !! " & strCrLf & "�_�h�N�L�k�Ұ� Call Center �\�� !!", vbInformation, "�T��"
'        Else
'            CreateObject("WScript.Shell").RegWrite ("HKCU\SocketInfo"), (ServerStr & "�" & PortStr & "�0")
'        End If
'    End If
    
    On Error Resume Next
    If rsChkStation.State = 1 Then rsChkStation.Close
    Set rsChkStation = Nothing
    If rsOnLine.State = 1 Then rsOnLine.Close
    Set rsOnLine = Nothing
  Exit Function
ChkErr:
    Static ErrTime As Integer
    ErrTime = ErrTime + 1
    If ErrTime > 3 Then
        ErrTime = 0
        Err.Clear
        Resume Next
    Else
        If Err.Number = 3021 Then
            Resume 0
            Err.Clear
        Else
            ErrSub Me.Name, "ChkComputerName"
        End If
    End If
End Function

Private Function ChkDupLogin() As Boolean
  On Error GoTo ChkErr
    Dim intCanDupLogin As Integer
    On Error Resume Next
    intCanDupLogin = gcnGi.Execute("SELECT DUPLOGIN FROM " & GetOwner & "SO041 WHERE ROWNUM=1").GetString
'   MsgBox Err.Number
    On Error GoTo ChkErr
    If Err.Number <> 0 Then
        Select Case Err.Number
               Case 3021
                    MsgBox "�Х��]�w [�t�η~�̰Ѽ���](SO041) ����� , �t�ΧY�N����!!", vbInformation, "�T��"
               Case -2147217900
                    MsgBox "�b�}�� [�t�η~�̰Ѽ���](SO041) ����Ʈɵo�Ϳ��~" & vbCrLf & "�䤣�� [DupLogin] ��� , �t�ΧY�N����!!", vbInformation, "�T��"
               Case Else
                    MsgBox "�b�}�� [�t�η~�̰Ѽ���](SO041) ����Ʈɵo�Ϳ��~ , �t�ΧY�N����!!", vbInformation, "�T��"
        End Select
        End
    End If
    If Val(gcnGi.Execute("SELECT COUNT(*) FROM " & GetOwner & "SO025 WHERE USERID='" & txtUserId & "' And StationName <> '" & gstrPcName & "'").GetString) > 0 Then
        If intCanDupLogin = 1 Then
            If Not Alfa Then MsgBox "���ϥΪ̤w�b�t�Τ��n�J!!", vbInformation, "�T��"
            ChkDupLogin = True
        Else
            MsgBox "���ϥΪ̤w�b�t�Τ��n�J!! �Фŭ��еn�J", vbInformation, "�T��"
            ChkDupLogin = False
            txtUserId.Text = ""
            txtPassword.Text = ""
            On Error Resume Next
            txtUserId.SetFocus
        End If
    Else
        ChkDupLogin = True
    End If
  Exit Function
ChkErr:
    ErrSub Me.Name, "ChkDupLogin"
End Function

' ��J����Ƥ��ŦX�ɡA�M����J����̪���J���
Public Sub WrongPwd()
  On Error GoTo ER
    txtUserId = ""
    txtPassword = ""
    txtUserId.SetFocus
  Exit Sub
ER:
   ErrSub Me.Name, "WrongPwd"
End Sub

Private Function ReadCDRINI() ' �� GICMIS2.INI Ū���]�w���
  On Error GoTo ChkErr
    Dim SysPath As String
    Dim strTmp As String
    Dim lngLen As Long
    Dim pstrDataSource As String
    Dim pstrUserId As String
    Dim pstrMDBPath As String
    Dim pstrPassword As String
    Dim cpCode As String

    SysPath = GetIniPath2
    
    If Dir(SysPath) = "" Then ' �p�G�٬O�䤣�� GICMIS.INI �A�h�����t��
        MsgBox "�t�����ҰѼƳ]�w���~�A�L�k�}��" & SysPath & " !", vbCritical, "���~"
        Unload Me
        End
    End If
    
    cpCode = ReadFROMINI(SysPath, "DEFAULT", "DEFAULTDB", False) ' �� GICMIS2.INI Ū�� �]�w��
    garyGi(14) = cpCode
    pstrDataSource = ReadFROMINI(SysPath, cpCode, "1", True)
    pstrUserId = ReadFROMINI(SysPath, cpCode, "2", True)
    pstrPassword = ReadFROMINI(SysPath, cpCode, "3", True)
    gstrStationCount = ReadFROMINI(SysPath, cpCode, "4", True) ' ���\�h�֤u�@���n�J
    
    If Len(pstrUserId) = 0 Or Len(pstrPassword) = 0 Or Len(gstrStationCount) = 0 Then
        MsgBox "�t �� �� �� �� �� �] �w �� �~ !!", , "���~�T��..!!"
        Unload Me ' �p�G ��ƨӷ� �� �ϥΪ̥N�X �ťաA
        End
    End If
   
    If Alfa Then
        GetGlobal
    Else
        garyGi(3) = "Provider=MSDAORA.1;" & _
                     IIf(pstrDataSource = "", "", "Data Source=" + pstrDataSource + ";") & _
                    "Password=" & pstrPassword + ";" & _
                    "User ID=" & pstrUserId + ";" & _
                    "Persist Security Info=True"
    End If
    Exit Function
ChkErr:
    If ErrHandle(, True, , "ReadINI") Then Resume 0
    Unload Me
    End
End Function

Private Function ReadINI2() ' �� MonitorCenter.INI Ū���]�w���
  On Error GoTo ChkErr
    Dim SysPath As String
    Dim strTmp As String
    Dim lngLen As Long
    Dim pstrDataSource As String
    Dim pstrUserId As String
    Dim pstrMDBPath As String
    Dim pstrPassword As String
    Dim cpCode As String

    SysPath = GetCDRIniPath
    
    If Dir(SysPath) = "" Then ' �p�G�٬O�䤣�� GICMIS.INI �A�h�����t��
        MsgBox "�ϥΪ̳]�w���~�A�L�k�}��" & SysPath & " !", vbCritical, "���~"
        Unload Me
        End
    End If
    
    txtUserId.Text = ReadFROMINI(SysPath, "LogIn", "Uid")
    txtPassword.Text = ReadFROMINI(SysPath, "LogIn", "Pwd")
'    gstrStationCount = ReadFROMINI(SysPath, cpCode, "4", True) ' ���\�h�֤u�@���n�J
    
    If Len(txtUserId.Text) = 0 Or Len(txtPassword.Text) = 0 Then
        MsgBox "�� �� �� �] �w �� �~ !!", , "���~�T��..!!"
        Unload Me ' �p�G ��ƨӷ� �� �ϥΪ̥N�X �ťաA
        End
    Else
        Call cmdOk_Click
    End If
   
    Exit Function
ChkErr:
    If ErrHandle(, True, , "ReadINI2") Then Resume 0
    Unload Me
    End
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = 27 Then cmdCancel_Click
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
    With Me
        .Move (Screen.Width - .Width) / 2, ((Screen.Height - .Height) / 2)
    End With
'   MakeTransparent Me.hwnd, 99
    'ExplodeForm Me, 24, &HFF0000
    rs.CursorLocation = adUseClient ' �� SO041 Ū���~�̦W�١B�{���W�ٻP����
    If OpenRecordset(rs, "SELECT SysName, PrgName, Version FROM " & GetOwner & "SO041 Order By SysId", gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False) <> giOK Then
        Call ErrHandle("�L�k��SO041���o�~�̦W�١B�{���W�ٻP����!�o�ӵ{���Y�N����!", False, 2, "Form_Initialize")
        End
        Exit Sub
    End If
    If rs.RecordCount < 1 Then
        Call ErrHandle("�Х��]�w�~�̦W�١B�{���W�ٻP����!", False, 2, "Form_Initialize")
        End
        Exit Sub
    End If
    lblCompName = rs("SysName").Value & ""
    lblSysName = rs("PrgName").Value & ""
    lblVersion = rs("Version").Value & ""
    rs.Close
    Set rs = Nothing
'    If Alfa Then txtUserId = "A": txtPassword = "A": cmdOk.Value = True
    Call ReadINI2
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub Form_Resize()
  On Error Resume Next
    imgExit.Top = 55
    imgBottom.Left = 0
    imgLeft.Top = 0
    imgRight.Top = 0
    imgTop.Top = 0
    imgLeft.Left = 0
    imgTopRight.Left = Me.ScaleWidth - imgTopRight.Width
    imgTopLeft.Left = 0
    imgTop.Width = Me.ScaleWidth
    imgLeft.Height = Me.ScaleHeight
    imgRight.Height = Me.ScaleHeight
    imgRight.Left = Me.ScaleWidth - imgRight.Width
    imgBottom.Top = Me.ScaleHeight - imgBottom.Height
    imgBottom.Width = Me.ScaleWidth
    imgExit.Left = Me.ScaleWidth - imgExit.Width - imgLeft.Width - 40
    imgMaxButton.Top = imgExit.Top
    imgMinButton.Top = imgExit.Top
    If imgMaxButton.Visible Then
        imgMaxButton.Left = imgExit.Left - imgMinButton.Width - 50
        imgMinButton.Left = imgExit.Left - imgMinButton.Width - imgMaxButton.Width - 100
    Else
        imgMinButton.Left = imgExit.Left - imgMinButton.Width - 50
    End If
    imgLeft.ZOrder
    imgRight.ZOrder
    imgBottom.ZOrder
    imgTopRight.ZOrder
    imgTopLeft.ZOrder
    imgExit.ZOrder
    lblTestShadow.Left = lblTest.Left + 10
    lblTestShadow.Top = lblTest.Top + 10
    lblTest_Change
    lblTest.ZOrder
    imgIcon.ZOrder
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
'    If Len(Environ("OS")) > 0 Then
'        Dim intLoop As Integer
'        MakeTransparent Me.hwnd, 255
'        For intLoop = 255 To 0 Step -15
'            MakeTransparent Me.hwnd, intLoop
'            DoEvents
'        Next
'    End If
    Set OpenDB = Nothing
    If gcnGi.State > 0 Then gcnGi.Close
    Set gcnGi = Nothing
    DoEvents
    ReleaseCOM Me
    End
End Sub

Private Sub txtPassword_GotFocus()
  On Error Resume Next
    objSelectAll txtPassword
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyReturn Then cmdOk_Click
  Exit Sub
ChkErr:
    ErrSub Me.Name, "txtPassword_KeyDown"
End Sub

Private Sub txtUserId_GotFocus()
  On Error Resume Next
    objSelectAll txtUserId
End Sub

Private Sub txtUserId_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If Len(txtPassword.Text) = 0 Then
        If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
    Else
        If KeyCode = vbKeyReturn Then cmdOk_Click
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "txtUserId_KeyDown"
End Sub

Private Sub imgExit_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgExit.Picture = imgDisabled(2).Picture
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgExit.Picture = imgEnabled(2).Picture
End Sub

Private Sub imgIcon_DblClick()
  On Error Resume Next
    If imgExit.Enabled Then imgExit_Click
End Sub

Private Sub imgMaxButton_Click()
  On Error Resume Next
    Select Case Me.WindowState
           Case 0
                Me.WindowState = 2
           Case 2
                Me.WindowState = 0
    End Select
End Sub

Private Sub imgMaxButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgMaxButton.Picture = imgDisabled(0).Picture
End Sub

Private Sub imgMaxButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgMaxButton.Picture = imgEnabled(0).Picture
End Sub

Private Sub imgMinButton_Click()
  On Error Resume Next
    Me.WindowState = 1
End Sub

Private Sub imgMinButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgMinButton.Picture = imgDisabled(1).Picture
End Sub

Private Sub imgMinButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgMinButton.Picture = imgEnabled(1).Picture
End Sub

Private Sub imgTop_DblClick()
  On Error GoTo ChkErr
    If imgMaxButton.Enabled And imgMaxButton.Visible Then
        imgMaxButton_Click
    End If
  Exit Sub
ChkErr:
    ErrHandle Me.Name, "imgTop_DblClick"
End Sub

Private Sub imgTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    If Button = 1 Then
        mouseX = X
        mouseY = Y
    End If
End Sub

Private Sub imgTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    If Button = 1 Then
        Me.Left = Me.Left + X - mouseX
        Me.Top = Me.Top + Y - mouseY
    End If
End Sub

Private Sub imgTopLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgTopLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgTopRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblTest_Change()
  On Error Resume Next
    lblTestShadow.Caption = lblTest.Caption
    SetWindowText Me.hwnd, lblTest.Caption
End Sub

Private Sub lblTest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblTest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    imgTop_MouseMove Button, Shift, X, Y
End Sub
