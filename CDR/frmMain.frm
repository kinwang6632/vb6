VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "CDRCenter"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   3525
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3525
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame fraAuto 
      Caption         =   "�۰ʤưѼ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   3405
      Begin VB.TextBox txtAutoRun 
         Alignment       =   1  '�a�k���
         Enabled         =   0   'False
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Text            =   "1"
         Top             =   390
         Width           =   585
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  '��u�T�w
         Caption         =   "�ݩR��.."
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2250
         TabIndex        =   3
         Top             =   420
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "������۰ʰ���"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   870
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Timer tmrRun 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   630
      Top             =   1410
   End
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   150
      Top             =   1410
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intAutoRunCut As Long
Private intAutoConnCut As Long
Private strConnectFile As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private RetSections() As String
Private lngErrCnt As Long
Private Const fraCaption As String = "�۰ʤưѼ�"
Private Const SendSmsCom As String = "SENDSMSCOM"
Private Const SendSmsTel As String = "SMSTEL"
Private Const SendSmsErrCount As String = "SENDSMSERR"
Private strSmsConnectString As String
Private strSmsTel As String
Private lngSmsErrCount As Long
Private Sub Form_Initialize()
  On Error GoTo ChkErr
    SetLocaleInfo GetSystemDefaultLCID, &H1F, "yyyy/MM/dd"
    gstrCurDir = CurPath
    If OpenDB Is Nothing Then Set OpenDB = CreateObject("giOpenDBObj.OpenDBObj")
    Call ReadCDRINI(Me) ' �� GICMIS2.INI Ū����T
    Call ReadSmsInfo
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Initialize")
End Sub
'Ū���ǰe��SMS����T
Private Sub ReadSmsInfo()
  On Error GoTo ChkErr
    Dim smsDataSource As String
    Dim smsUserId As String
    Dim smsPassword As String
    Dim aIniFile As String
    Dim strSplit() As String
    Dim aTel As String
    Dim i As Long
    strSmsConnectString = Empty
    strSmsTel = Empty
    aIniFile = GetCDRIniPath
    smsDataSource = ReadFROMINI(aIniFile, SendSmsCom, "1", True)
    smsUserId = ReadFROMINI(aIniFile, SendSmsCom, "2", True)
    smsPassword = ReadFROMINI(aIniFile, SendSmsCom, "3", True)
    
    If (Len(smsDataSource) > 0) And (Len(smsUserId) > 0) And (Len(smsPassword) > 0) Then
    
        strSmsConnectString = "Provider=MSDAORA.1;" & _
                         IIf(smsDataSource = "", "", "Data Source=" + smsDataSource + ";") & _
                        "Password=" & smsPassword + ";" & _
                        "User ID=" & smsUserId + ";" & _
                        "Persist Security Info=True"
                        
        lngSmsErrCount = Val(ReadFROMINI(aIniFile, SendSmsErrCount, "Count") & "")
        If lngSmsErrCount <= 0 Then
            lngSmsErrCount = 1
        End If
        
        strSplit = Split(ReadSectionAllKeys(aIniFile, SendSmsTel), ",")
        For i = LBound(strSplit) To UBound(strSplit)
            aTel = ReadFROMINI(aIniFile, SendSmsTel, strSplit(i), False)
            If Len(aTel) > 0 Then
                If strSmsTel = Empty Then
                    strSmsTel = aTel
                Else
                    strSmsTel = strSmsTel & "," & aTel
                End If
            End If
        Next i
    End If
    
    'strSmsTel
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "ReadSmsInfo")
End Sub
Private Sub ReadAllSections(ByVal aIniFile As String)
  On Error Resume Next
    Dim lngLength As Long
    Dim Ret As String
    
    Dim i As Long
    Ret = String(1024, 0)
    lngLength = GetPrivateProfileSectionNames(Ret, 1024, aIniFile)
    Ret = Left$(Ret, lngLength)
    If lngLength > 1 Then
        RetSections = Split(Ret, vbNullChar)
    End If
End Sub
Private Function ReadSectionAllKeys(ByVal aIniFile As String, ByVal aSection As String) As String
  On Error GoTo ChkErr
    Dim strKeys As String
    Dim strSplit() As String
    Dim i As Long
    strKeys = ReadFROMINI(aIniFile, aSection, vbNullString)
    strKeys = Trim(Replace(strKeys, vbCrLf, ""))
    strSplit = Split(strKeys, Chr(0))
    strKeys = Empty
    For i = LBound(strSplit) To UBound(strSplit)
        If strSplit(i) <> Empty Then
            If strKeys = Empty Then
                strKeys = strSplit(i)
            Else
                strKeys = strKeys & "," & strSplit(i)
            End If
        End If
    Next
    ReadSectionAllKeys = strKeys
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "ReadAllKeys")
End Function
Private Function ISSectionExists(ByVal aSection As String) As Boolean
  On Error GoTo ChkErr
    Dim i As Long
    
    For i = LBound(RetSections) To UBound(RetSections)
        If Trim(RetSections(i)) = Trim(aSection) Then
            ISSectionExists = True
            Exit For
        End If
    Next
    Exit Function
ChkErr:
  ISSectionExists = False
End Function
Private Sub Form_Load()
  On Error GoTo ChkErr
    Dim obj As Object
    Dim smsDataSource As String
    Dim smsUserId As String
    Dim smsPassword As String
    
    If Alfa Then
        MsgBox "Alfa=True"
    End If
   
    If Not Alfa Then
        Set obj = CreateObject("csLockProcess.EnableProcess")
        If Not obj.EnableProcess(GetCDRIniPath, "Lock", "Over") Then
            MsgBox "�{�����X�k�I�Ь��}�լ��", vbCritical, "ĵ�i"
            End
        End If
    End If
    
    
'    smsDataSource = ReadFROMINI(GetCDRIniPath, SendSmsCom, "1", True)
'    smsUserId = ReadFROMINI(GetCDRIniPath, SendSmsCom, "2", True)
'    smsPassword = ReadFROMINI(GetCDRIniPath, SendSmsCom, "2", True)
'    If (Len(Trim(smsDataSource)) > 0) And (Len(Trim(smsUserId)) > 0) And (Len(Trim(smsPassword)) > 0) Then
'        SMSConnectString = "Provider=MSDAORA.1;" & _
'                     IIf(smsDataSource = "", "", "Data Source=" + smsDataSource + ";") & _
'                    "Password=" & smsPassword + ";" & _
'                    "User ID=" & smsUserId + ";" & _
'                    "Persist Security Info=True"
'    End If
    
    If Alfa Or blnShowSecond Then
        
        tmrRun.Interval = 1000
        'tmrConnect.Interval = 1000
    Else
        tmrRun.Interval = 60000
        'tmrConnect.Interval = 1000
    End If
    If IsConnect(gcnGi) Then
        Call DataInitial
        tmrRun.Enabled = True
        'tmrConnect.Enabled = False
    Else
        tmrRun.Enabled = True
        'tmrConnect.Enabled = True
    End If
    
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
    Call CloseRecordset(rsSO033Vod)
    Call CloseRecordset(rsSO555C)
    'Call CloseRecordset(rsReadFile)
    'WriteCDRParaIni
    End
End Sub

Private Sub tmrRun_Timer()
  On Error Resume Next
    
'     garyGi(3) = "Provider=MSDAORA.1;" & _
'                     IIf(pstrDataSource = "", "", "Data Source=" + pstrDataSource + ";") & _
'                    "Password=" & pstrPassword + ";" & _
'                    "User ID=" & pstrUserId + ";" & _
'                    "Persist Security Info=True"
    
'    pstrDataSource = ReadFROMINI(SysPath, cpCode, "1", True)
'    pstrUserId = ReadFROMINI(SysPath, cpCode, "2", True)
'    pstrPassword = ReadFROMINI(SysPath, cpCode, "3", True)
'    gstrStationCount = ReadFROMINI(SysPath, cpCode, "4", True) ' ���\�h�֤u�@���n�J
'
    Me.Enabled = False
    intAutoRunCut = intAutoRunCut - 1
    If intAutoRunCut <= 0 Then
        intAutoRunCut = intAutoCDR
        If Not IsConnect(gcnGi) Then
            Me.Enabled = True
            Exit Sub
        Else
            Call ExeRun      '�}�l����
        End If
    End If
    Me.Enabled = True
End Sub
Private Sub EXERun2()
  On Error GoTo ChkErr
    tmrRun.Enabled = False
    Dim obj As Object
    Set obj = CreateObject("csCallExeCDR.clsExeCmd")
    Dim i As Long
    obj.SetWaitTime = 60
    Set obj.uCn = gcnGi
    obj.SysPath = GetErrFile
    'obj.CDRIniFile = GetCDRIniPath
    obj.ExeRunTime = GetRightNow
    'obj.uErrCount = lngErrExitCount
    For i = LBound(varOwner) To UBound(varOwner)
        Call SetPubOwner(i)
        obj.ugaryGi = Join(garyGi, Chr(9))
        obj.uOwnerName = GetOwner
        If ISSectionExists(gCompCode) Then
            If Not IsConnect(gcnGi) Then
                lngErrCnt = lngErrCnt + 1
                If lngErrCnt >= lngErrExitCount And lngErrExitCount > 0 Then
                    Call ExitApp
                End If
                GoTo ChkErr
            End If
            Dim rs As New ADODB.Recordset
            Dim strQry As String
            strQry = "SELECT A.*,A.ROWID FROM " & GetOwner & "SO004 A WHERE custid=658942 and seqno='201002040268282'"
            If Not GetRS(rs, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
            If obj.SendZ2Cmd(Now, Now, rs, strVODCmdSeqNo) Then
                
                If Not obj.ExeAllCommit(strVODCmdSeqNo, _
                                        , True) Then
                    
                    Exit Sub
                
                End If
            Else
                lngErrCnt = lngErrCnt + 1
                If lngErrCnt >= lngErrExitCount And lngErrExitCount > 0 Then
                    Call ExitApp
                End If
            End If
        End If
    Next
    Set obj = Nothing
    If Alfa = True Or blnShowSecond Then
        tmrRun.Interval = 6000              '�Ҧ������q�O���]���NTimer���]��1����
    Else
        tmrRun.Interval = 60000
    End If
    'intAutoRunCut = intAutoRunCut - DateDiff("n", Now, t1)  '�p��g�L�X�����F
    If intAutoRunCut <= 0 Then tmrRun.Interval = 100    '�p�G����ɶ��w�g�W�LTimer,�h�ߧY����
    tmrRun.Enabled = True
    Exit Sub
ChkErr:
  Set obj = Nothing
  tmrRun.Enabled = True
  Call ErrCDRSub(Me.Name, "ExeRun2")
End Sub
Private Sub ExeRun()
  On Error GoTo ChkErr
    lblStatus.Caption = "���椤.."
    lblStatus.BackColor = vbRed
    tmrRun.Enabled = False
    Dim obj As Object
    Set obj = CreateObject("csCallExeCDR.clsExeCmd")
    Dim i As Long
    Dim t1 As Date
    t1 = Now
    obj.SetWaitTime = lngCmdTimeOut
    Set obj.uCn = gcnGi
    obj.SysPath = GetErrFile
    obj.CDRIniFile = GetCDRIniPath
    obj.ExeRunTime = GetRightNow
    obj.uErrCount = lngErrExitCount
    obj.uSmsConnectString = strSmsConnectString
    obj.uSmsTels = strSmsTel
    obj.uSmsErrCount = lngSmsErrCount
    For i = LBound(varOwner) To UBound(varOwner)
        Call SetPubOwner(i)
        obj.ugaryGi = Join(garyGi, Chr(9))
        obj.uOwnerName = GetOwner
        If ISSectionExists(gCompCode) Then
            If Not IsConnect(gcnGi) Then
                lngErrCnt = lngErrCnt + 1
                If lngErrCnt >= lngErrExitCount And lngErrExitCount > 0 Then
                    Call ExitApp
                End If
                GoTo ChkErr
            End If
            If obj.SendZ3Cmd("", strVODCmdSeqNo, True) Then
                
                If Not obj.ExeOneCommit(strVODCmdSeqNo, lngErrCnt, _
                                        MAXIncurredDate, True, True, True, True) Then
                            
                    If lngErrCnt >= lngErrExitCount And lngErrExitCount > 0 Then
                        
                        Call ExitApp
            
                    End If
                    
                Else
                    If lngErrCnt >= lngErrExitCount And lngErrExitCount > 0 Then
                        Call ExitApp
                    End If
                End If
            Else
                lngErrCnt = lngErrCnt + 1
                If lngErrCnt >= lngErrExitCount And lngErrExitCount > 0 Then
                    Call ExitApp
                End If
            End If
        End If
    Next
    Set obj = Nothing
    If Alfa = True Or blnShowSecond Then
        tmrRun.Interval = 6000              '�Ҧ������q�O���]���NTimer���]��1����
    Else
        tmrRun.Interval = 60000
    End If
    intAutoRunCut = intAutoRunCut - DateDiff("n", Now, t1)  '�p��g�L�X�����F
    If intAutoRunCut <= 0 Then tmrRun.Interval = 100    '�p�G����ɶ��w�g�W�LTimer,�h�ߧY����
    lblStatus.Caption = "�ݩR��.."
    lblStatus.BackColor = vbGreen
    tmrRun.Enabled = True
    Exit Sub
ChkErr:
  Set obj = Nothing
  tmrRun.Enabled = True
  Call ErrCDRSub(Me.Name, "ExeRun")
End Sub

Private Sub ExitApp()
  On Error Resume Next
    MsgBox "���~�w�F [ " & lngErrExitCount & " ] ��" & vbCrLf & _
                            "�Ьd�� " & GetErrFile & "SYSERR.TXT ��", vbInformation, "�T��"
    End
End Sub


'���q�\��²�z�p�U
'�ǰeZ3�R�O�ܱ���x,�H�ἴ��SO555C�����
'�HSO555C����Ʒs�W��SO033VOD,�A��sSO004G���I�ƭ�,�H��s���I�ƭȦA�s�W��SO182
'��²��a..�藍��
'Private Sub ExeRun()
'  On Error Resume Next
'    Dim t1 As Date
'    Dim i As Long
'    Dim blnFinish As Boolean
'    Dim blnBeginTrans As Boolean
'    Dim blnAllFinish As Boolean
'    t1 = Now
'
'    tmrRun.Enabled = False
'    For i = LBound(varOwner) To UBound(varOwner)
'        blnFinish = False
'        blnBeginTrans = False
'        Call SetPubOwner(i)
'        If Not IsConnect(gcnGi) Then GoTo ChkErr
'
'        strRunTime = GetRightNow
'        strSendCmdTime = GetSendTime    '�qini�ɨ��X�n�ǰe�R�O���_�l���
'        If strSendCmdTime <> "" Then
'            strSendCmdTime = Format(strSendCmdTime, "yyyymmddhhmmss")
'            gdtSendTime.SetValue strSendCmdTime
'        Else
'            GoTo lNext
'        End If
'        If Not SendZ3Cmd(gdtSendTime.GetValue(True)) Then GoTo lNext    '�ǰe�R�O�ܱ���x
'        If Not GetSO555C(strVODCmdSeqNo, rsSO555C) Then GoTo lNext        '����SO555C�����
'        MAXIncurredDate = gdtSendTime.GetValue(True)
'        If Not rsSO555C.EOF Then
'            rsSO555C.MoveFirst
'            blnAllFinish = True
'            Do While Not rsSO555C.EOF
'                DoEvents
'                If CDate(rsSO555C("BIncurredDate")) > MAXIncurredDate Then
'                    MAXIncurredDate = rsSO555C("BIncurredDate") & ""
'                End If
'                rsSO555C.ActiveConnection = Nothing
'                If Not IsConnect(gcnGi) Then GoTo ChkErr
'                Dim blnChkSO033VOD As Boolean
'                blnChkSO033VOD = False
'                gcnGi.BeginTrans
'                blnBeginTrans = True
'                blnFinish = False
'                If Not InsSO033VOD(rsSO555C, blnChkSO033VOD) Then        '�s�W��Ʀ�SO033VOD
'                    If Not blnChkSO033VOD Then GoTo lDo     '�p�G�O�]��SO033VOD���Ƥ]�n�⦨�\
'                End If
'                If Not blnChkSO033VOD Then                  'SO033VOD�����ƴN���n�s�WSO182��UPD SO004G�F
'                    Dim aFrequencyType As String
'                    aFrequencyType = UCase(rsSO555C("FrequencyType") & "")
'                    If (aFrequencyType = "IMP") Or (aFrequencyType) = "MUL" Then
'                        '��sSO004G���I�ƭ�,�÷s�W��Ʀ�SO182
'                        If Not UpdSO004G(rsSO555C("VODACCOUNTID") & "", _
'                                        rsSO555C("VALUE") & "", True, True, rsSO033Vod) Then GoTo lDo
'                    End If
'
'                End If
'                blnFinish = True
'                gcnGi.CommitTrans
'                rsSO555C.MoveNext
'lDo:
'                If Not blnFinish And blnBeginTrans Then
'                    gcnGi.RollbackTrans
'                    blnAllFinish = False    '�p�G���@������,�h���O��False,�����\�g�^INI��
'                End If
'            Loop
'            If blnAllFinish Then Call WriteCDRParaIni        '�������\�~�N�̤j���ɶ��g�JINI��,�ñN��l�ƭȧאּ 1
'        End If
'lNext:
'    Next i
'    If Alfa = True Then
'        tmrRun.Interval = 6000              '�NTimer���]��1����
'    Else
'        tmrRun.Interval = 60000
'    End If
'88:
'    intAutoRunCut = intAutoRunCut - DateDiff("n", Now, t1)
'    If intAutoRunCut <= 0 Then tmrRun.Interval = 100    '�p�G����ɶ��w�g�W�LTimer,�h�ߧY����
'    tmrRun.Enabled = True
'    Exit Sub
'ChkErr:
'    Call ErrCDRSub(Me.Name, "ExeRun")
'    intAutoRunCut = intAutoRunCut - DateDiff("n", Now, t1)
'    If intAutoRunCut <= 0 Then tmrRun.Interval = 100
'    tmrRun.Enabled = True
'End Sub

Private Sub DataInitial()
  On Error Resume Next
    Call GetCDRParaIni  '���XINI�ɪ��]�w��
    Call ReadAllSections(GetCDRIniPath)
    'gdtSendTime.ShowSecond
    Call SetCompCodeAry
    txtAutoRun.Text = intAutoCDR
    intAutoRunCut = intAutoCDR
    If strUpdEn <> "" Then
        garyGi(0) = strUpdEn
        garyGi(1) = strUpdEn
    End If
    If blnShowSecond Then tmrRun.Interval = 1000
    If strAppTitle <> "" Then
        Me.Caption = strAppTitle
    End If
End Sub

