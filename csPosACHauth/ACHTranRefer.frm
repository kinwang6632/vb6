VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmACHTranRefer 
   Caption         =   "ACH��b���ڸ�Ʋ���[ACHTranRefer]"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8970
   Icon            =   "ACHTranRefer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   8970
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   3270
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�}�l"
      Height          =   375
      Left            =   165
      TabIndex        =   10
      Top             =   3270
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3060
      Left            =   165
      TabIndex        =   12
      Top             =   105
      Width           =   8595
      Begin VB.TextBox txtPutAccount 
         Height          =   300
         Left            =   5790
         MaxLength       =   14
         TabIndex        =   5
         Top             =   1140
         Width           =   1830
      End
      Begin VB.TextBox txtGotSecId 
         Height          =   300
         Left            =   2025
         MaxLength       =   7
         TabIndex        =   1
         Top             =   780
         Width           =   1830
      End
      Begin VB.Frame frmData 
         Caption         =   "����ɦ�m"
         Height          =   1275
         HelpContextID   =   2
         Left            =   210
         TabIndex        =   13
         Top             =   1650
         Width           =   8160
         Begin VB.TextBox txtErrPath 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3015
            TabIndex        =   8
            Top             =   765
            Width           =   4095
         End
         Begin VB.CommandButton cmdDataPath 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   14.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7320
            TabIndex        =   7
            Top             =   315
            Width           =   375
         End
         Begin VB.CommandButton cmdErrPath 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   14.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7335
            TabIndex        =   9
            Top             =   780
            Width           =   375
         End
         Begin VB.TextBox txtDataPath 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3015
            TabIndex        =   6
            ToolTipText     =   "�п�J�r�ɤ����|���ɦW�I"
            Top             =   300
            Width           =   4095
         End
         Begin VB.Label lblErrpath 
            AutoSize        =   -1  'True
            Caption         =   "���D�Ѧ��ɦ�m�]���| + �W�١^"
            Height          =   195
            Left            =   510
            TabIndex        =   15
            Top             =   810
            Width           =   2460
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "����ɦ�m�]���| +�W��)"
            Height          =   180
            Left            =   930
            TabIndex        =   14
            Top             =   375
            Width           =   1995
         End
      End
      Begin VB.TextBox txtInvoiceId 
         Height          =   345
         Left            =   2025
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1140
         Width           =   1830
      End
      Begin VB.TextBox txtPutBankId 
         Height          =   300
         Left            =   5790
         MaxLength       =   7
         TabIndex        =   4
         Top             =   780
         Width           =   1830
      End
      Begin VB.TextBox txtSendSpcId 
         Height          =   300
         Left            =   2025
         MaxLength       =   7
         TabIndex        =   0
         Top             =   435
         Width           =   1830
      End
      Begin Gi_Date.GiDate gdaHandleDate 
         Height          =   345
         Left            =   5790
         TabIndex        =   3
         Top             =   435
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  '�a�k���
         Caption         =   "�������N��"
         Height          =   285
         Left            =   510
         TabIndex        =   21
         Top             =   870
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   1  '�a�k���
         Caption         =   "�o�ʪ̱b��"
         Height          =   210
         Left            =   4320
         TabIndex        =   20
         Top             =   1170
         Width           =   1230
      End
      Begin VB.Label Label3 
         Alignment       =   1  '�a�k���
         Caption         =   "�o�ʪ̲Τ@�s��"
         Height          =   210
         Left            =   525
         TabIndex        =   19
         Top             =   1245
         Width           =   1410
      End
      Begin VB.Label Label8 
         Alignment       =   1  '�a�k���
         Caption         =   "�B�z���"
         Height          =   285
         Left            =   4485
         TabIndex        =   18
         Top             =   435
         Width           =   1050
      End
      Begin VB.Label Label5 
         Alignment       =   1  '�a�k���
         Caption         =   "���X��N��"
         Height          =   285
         Left            =   4290
         TabIndex        =   17
         Top             =   780
         Width           =   1260
      End
      Begin VB.Label Label7 
         Alignment       =   1  '�a�k���
         Caption         =   "�o�e���N��"
         Height          =   285
         Left            =   210
         TabIndex        =   16
         Top             =   510
         Width           =   1725
      End
   End
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   7755
      Top             =   570
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmACHTranRefer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strPath As String
Private strChoose As String
Private strCompCode As String
Private strACHTNo As String
Private strBankHand As String
Private strAccNId As String
Private intPara24 As String
Private strPrgName As String
Private strBankId2 As String
Private strErrPath As String
Private strUpdEn As String
Private strUpdName As String
Private rsTmp As New ADODB.Recordset
Dim intAmount As Double
Private DataPath As TextStream
Private ErrPath As TextStream

Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdCancel")
End Sub

Private Sub cmdDataPath_Click()
        On Error GoTo ChkErr
        With cnd1
                .FileName = txtDataPath.Text
                .Filter = "��r�� (*.txt;*.dat)|*.txt;*.dat|�Ҧ��ɮ� (*.*)|*.*"
                .InitDir = ReadGICMIS1("ErrLogPath")
                .Action = 1
                txtDataPath = .FileName
        End With
    Exit Sub
    
ChkErr:
    Call ErrSub(Me.Name, "cmdDataPath_Click")
End Sub


Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
Dim strErrMsg As String

    IsDataOk = False
    If gdaHandleDate.GetValue = "" Then strErrMsg = "��X���": gdaHandleDate.SetFocus: GoTo Warning
    If txtDataPath.Text = "" Then strErrMsg = "����ɦ�m": txtDataPath.SetFocus: GoTo Warning
    If txtErrPath.Text = "" Then strErrMsg = "����ɦ�m": txtErrPath.SetFocus: GoTo Warning
    IsDataOk = True
    Exit Function
Warning:
    Call MsgMustBe(strErrMsg)
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function


Private Sub cmdErrPath_Click()
        On Error GoTo ChkErr
        With cnd1
                .FileName = txtDataPath.Text
                .Filter = "��r�� (*.txt;*.dat)|*.txt;*.dat|�Ҧ��ɮ� (*.*)|*.*"
                .InitDir = ReadGICMIS1("ErrLogPath")
                .Action = 1
                txtErrPath = .FileName
        End With
    Exit Sub
    
ChkErr:
    Call ErrSub(Me.Name, "cmdErrPath_Click")
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ChkErr
    Dim S As String
    Dim Z As String
    If Not IsDataOk Then Exit Sub
    S = Left(txtDataPath, InStr(txtDataPath, "\"))
    If S = "C:\" Or S = "c:\" Then
    Else
        If Not ChkDirExist(S) Then MsgBox "���| " & S & " ���s�b!", vbExclamation: Exit Sub
    End If
    Z = Left(txtDataPath, InStr(txtErrPath, "\"))
    If Z = "C:\" Or Z = "c:\" Then
    Else
        If Not ChkDirExist(Z) Then MsgBox "���| " & Z & " ���s�b!", vbExclamation: Exit Sub
    End If
    If OpenFile(DataPath, txtDataPath, 0, True) = False Then Exit Sub
    If OpenFile(ErrPath, txtErrPath, 1, True) = False Then Exit Sub
'    If OpenFile(txtDataPath, True) = False Then Exit Sub

    Call ScrToRcd
    
    If Not BeginTran Then
            CloseFS
            Screen.MousePointer = vbDefault
            cmdOK.Enabled = True: cmdCancel.Enabled = True
            Unload Me
            Exit Sub
    End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdOk_Click")
End Sub
Private Function GetRsTmp(ByRef rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim strFrom  As String, strWhere As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset

       'EMC
'        If Not GetRS(rsTmp, "Select A.BankCode as BankCode,A.BankName as BankName,A.RealDate " & _
'          " as TranDate,A.AccountNo,A.ShouldAmt as ShouldAmt,A.BillNo as BillNo_Old,A.BillNo as BillNo_New From " & strOwnerName & "So033 A Where RowId=''") Then Exit Function
            strFrom = " " & strOwnerName & "SO033 A "
                If strChoose <> "" Then
                If InStr(1, strChoose, "C.") Then
                    strFrom = strFrom & "," & strOwnerName & "So014 C "
                    strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.AddrNo=C.AddrNo "
                End If
                If InStr(1, strChoose, "E.") Then
                    strFrom = strFrom & "," & strOwnerName & "So002 E "
                    strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.ServiceType = E.ServiceType And A.CustId=E.CustId And A.CompCode=E.CompCode "
                End If
            End If
            strWhere = strWhere & IIf(strWhere = "", "", " And ") & strChoose & " And A.CancelFlag=0 And A.UCCode is Not Null And A.ShouldAmt > 0"
            strSQL = "SELECT " & GetUseIndexStr("So033", "ShouldDate") & " A.MediaBillNo BillNo,A.AccountNo,A.BankCode,A.AddrNo, Sum(ShouldAmt) ShouldAmt  From " & _
                        strFrom & " Where " & _
                        strWhere & " Group By A.MediaBillNo,A.AccountNo,A.BankCode,A.AddrNo "
         
        If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
'        strReturnSQL = rs.Source
    GetRsTmp = True
    Exit Function
ChkErr:
     ErrSub Me.Name, "GetRsTmp"
End Function

Private Function GetRealDateTran(strRealDate As String) As String
    On Error Resume Next
        GetRealDateTran = Val(Replace(strRealDate, "/", "")) - 19110000
End Function

Private Function InsertHead(rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim intReserve As Integer
    Dim strData As Variant
        '�ϧO�X(���O) , ��ƥN��, �B�z���, �B�z�ɶ�, �o�e���N��, �������N��, �ƥ�
        strData = ""
        strData = strData & "BOF" & "ACHP01"
        '�����O(1-3)��ƥN��(4-9)
        strData = strData & GetString(GetRealDateTran(gdaHandleDate.Text), 8, giRight, True)
        '�B�z���(10-17)
        strData = strData & GetString(Format(Time, "HHMMSS"), 6, giLeft)
        '�B�z�ɶ�(18-23)
        strData = strData & GetString(txtSendSpcId, 7, giLeft)
        '�o�e���N��(24-30)
        strData = strData & GetString(txtGotSecId, 7, giLeft)
        '�������N��(31-37)
        strData = strData & GetString("", 123, giLeft)
        '�ƥ�(38-160)
        WriteLineData strData, DataPath
        InsertHead = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "InsertHead"

End Function

Private Function InsertFinal(rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim intReserve As Integer
    Dim strData As Variant
        strData = "EOF" & "ACHP01"
        '�����O(1-3) , ��ƥN��(4-9)
        strData = strData & GetString(GetRealDateTran(gdaHandleDate.Text), 8, giRight, True)
        '�B�z���(9-17)
        strData = strData & GetString(txtSendSpcId, 7, giLeft)
        '�o�e���N��(18-24)
        strData = strData & GetString(txtGotSecId, 7, giLeft)
        '�������N��(25-31)
        strData = strData & GetString(rs.RecordCount, 8, giRight, True)
        '�`����(32-39)
        strData = strData & GetString(intAmount, 14, giRight, True)
        strData = strData & GetString("", 2, , True)
        '�`���B(40-55)
        strData = strData & GetString("", 105, giLeft)
        '�ƥ�(56-160)
        
        WriteLineData strData, DataPath
        InsertFinal = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "InsertFinal"

End Function

Private Function InsertDetail(rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim strData As Variant
    Dim rsSo106 As New ADODB.Recordset
    Dim strSQL As String
    Dim intSeq As Integer
    intAmount = 0
    If rs.RecordCount > 0 Then rs.MoveFirst
'        On Error Resume Next
'    cnn.BeginTrans
    Do While Not rs.EOF
        strData = "N" & "SD"
        '������A(1-1) ������O(2-3)
        strData = strData & GetString(strACHTNo, 3, giLeft)
        '����N��(4-6)
        intSeq = intSeq + 1
        strData = strData & GetString(intSeq, 6, giLeft, True)
        '����Ǹ�(7-12)
        strData = strData & GetString(txtPutBankId, 7, giLeft)
        '���X��N��(13-19)
        strData = strData & GetString(txtPutAccount, 14, giRight, True)
        '�o�ʪ̱b��(20-33)
'        strData = strData & GetString(Abs(rs("RealAmt")), 12, giright, True)
'        strData = strData & GetString("", 2, , True)
        strBankId2 = GetRsValue("Select BankId2 From " & GetOwner & "CD018 Where CodeNo='" & rs("BankCode") & "" & "'") & ""
        strData = strData & GetString(strBankId2, 7, giLeft)
        '���^��N��(34-40)
        strData = strData & GetString(rs("AccountNo") & "", 14, giLeft)
        '�����̱b��(40-54)
        strData = strData & GetString(rs("ShouldAmt") & "", 10, giRight, True)
        '���B(55-64)
        strData = strData & GetString("", 2, giLeft)
        '�h��z�ѥN�� (65-66)
        strData = strData & "B" & GetString(txtInvoiceId, 10, giLeft)
        '���ܥ洫����(67-67)�o�ʪ̲Τ@�s��(68-77)
        strSQL = "SELECT AccountNameID,ACHCustId FROM " & GetOwner & "SO106 WHERE AccountID='" & rs("AccountNo") & _
                 "' And Bankcode='" & rs("BankCode") & "' And ACHCustId='" & strACHTNo & "' and AuthorizeStatus=1"
        If Not GetRS(rsSo106, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
        strAccNId = rsSo106("AccountNameID") & ""
        strData = strData & GetString(strAccNId, 10, giLeft)
        '�����̲Τ@�s��(78-87)
        strData = strData & GetString("", 21, giLeft)
        '�W���W�d���q�N��(88-93)�h�󥲭n���(94-108)�X�椣���L�X
        strData = strData & GetString(rsSo106("ACHCustid") & "", 20, giLeft)
        'SO106ACHCustid(109-128)
        strData = strData & GetString(rs("BillNo") & "", 24, giLeft)
        '�C��渹(128-147)
        strData = strData & GetString("", 8, giLeft)
        '�ƥ�(153-160)
        WriteLineData strData, DataPath

        intAmount = intAmount + Abs(rs("ShouldAmt"))
        rs.MoveNext
        DoEvents
    Loop
     InsertDetail = True
  Exit Function
ChkErr:
'    cnn.RollbackTrans
    ErrSub Me.Name, "InsertDetail"
End Function

Private Function BeginTran() As Boolean
    On Error GoTo ChkErr
'    Dim rsTmp As New ADODB.Recordset
    Dim strBankId As String, strSQL As String
    Dim strOldBillNo As String
    Dim lngCountCustId As Long, lngErrCount As Long
    Dim lngPara24 As Long
    Dim lngTime As Long
    Dim rsUpd As New ADODB.Recordset
        lngTime = Timer
'        �����
        Screen.MousePointer = vbHourglass
        cmdOK.Enabled = False: cmdCancel.Enabled = False
        If Not GetRsTmp(rsTmp) Then Exit Function

        If rsTmp.RecordCount = 0 Then
            MsgBox "�L��ơI�Э��s�ާ@�I", vbInformation, "�T���I"
            GoTo KillFile
        Else
            If Not InsertHead(rsTmp) Then GoTo KillFile
            If Not InsertDetail(rsTmp) Then GoTo KillFile
            If Not InsertFinal(rsTmp) Then GoTo KillFile
            MsgBox "�w������Ƶ��Ʀ@" & rsTmp.RecordCount & "��," & vbCrLf & vbCrLf & _
           "���D���Ʀ@" & lngErrCount & "��," & vbCrLf & vbCrLf & _
           "�@��O:" & (Timer - lngTime) \ 1 & "��"
'            msgResult rsTmp.RecordCount, lngErrCount, lngTime       '��ܰ��浲�G
        End If
        Screen.MousePointer = vbDefault
        cmdOK.Enabled = True: cmdCancel.Enabled = True
        Call CloseFS
        BeginTran = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
    Exit Function
KillFile:
'    �R���ɮ�..
    CloseFS
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "BeginTran")
End Function
Public Sub CloseFS()
    On Error Resume Next
    DataPath.Close
    ErrPath.Close
    Set fso = Nothing
End Sub
Private Sub ScrToRcd()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
        Set LogFile = fso.CreateTextFile(strErrPath & "\" & strPrgName & "Out.log", True)
        With LogFile
                .WriteLine (txtSendSpcId.Text)         '�o�e���N��
                .WriteLine (txtGotSecId.Text)          '�������N��
                .WriteLine (txtInvoiceId.Text)         '�o�ʪ̲Τ@�s��
                .WriteLine (gdaHandleDate.Text)        '�B�z���
                .WriteLine (txtPutBankId.Text)         '���X��N��
                .WriteLine (txtPutAccount.Text)        '�o�ʪ̱b��
                .WriteLine (txtDataPath.Text)          '����ɦ�m
                .WriteLine (txtErrPath.Text)           '���D�ɦ�m
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF2 Then Call cmdOk_Click: KeyCode = 0
    If KeyCode = vbKeyEscape Then Unload Me: KeyCode = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Screen.MousePointer = vbDefault
      Call InitData
      RcdToScr
End Sub

Private Sub RcdToScr()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    If Dir(strPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
            On Error Resume Next
            Set LogFile = fso.OpenTextFile(strErrPath & "\" & strPrgName & "Out.log", ForReading)
            With LogFile
                    If Not .AtEndOfStream Then txtSendSpcId.Text = .ReadLine & ""
                       '�o�e���N��
                    If Not .AtEndOfStream Then txtGotSecId.Text = .ReadLine & ""
                       '�������N��
                    If Not .AtEndOfStream Then txtInvoiceId.Text = .ReadLine & ""
                       '�o�ʪ̲Τ@�s��
                    If Not .AtEndOfStream Then gdaHandleDate.Text = .ReadLine & ""
                       '�B�z���
                    If Not .AtEndOfStream Then txtPutBankId.Text = .ReadLine & ""
                       '���X��N��
                    If Not .AtEndOfStream Then txtPutAccount.Text = .ReadLine & ""
                        '�o�ʪ̱b��
                    If Not .AtEndOfStream Then txtDataPath.Text = .ReadLine & ""
                       '����ɦ�m
                    If Not .AtEndOfStream Then txtErrPath.Text = .ReadLine & ""
                       '���D�ɦ�m
            End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Private Function InitData() As Boolean
    On Error GoTo ChkErr
        With frmSO3277A
             strChoose = .strChoose
             strCompCode = .gilCompCode.GetCodeNo
             strACHTNo = .txtACHTNo
'             strBankHand = .gilBankHand.GetCodeNo
'             intPara24 = .intPara24
             strPrgName = .strPrgName
        End With
        gdaHandleDate.SetValue (RightNow)
        strUpdEn = garyGi(0)
        strUpdName = garyGi(1)
        strErrPath = ReadGICMIS1("ErrLogPath")
        txtInvoiceId = GetRsValue("Select InvoiceId From " & strOwnerName & "So041") & ""
        InitData = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "InitData")
End Function



