VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUnion 
   BorderStyle     =   1  '��u�T�w
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUnion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9075
   StartUpPosition =   1  '���ݵ�������
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   6840
      Top             =   4245
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   4020
      Left            =   195
      TabIndex        =   14
      Top             =   135
      Width           =   8730
      Begin VB.ComboBox cboMonthDay 
         Height          =   315
         ItemData        =   "frmUnion.frx":0442
         Left            =   7590
         List            =   "frmUnion.frx":049A
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   29
         Top             =   330
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CheckBox chkDuteDate 
         Caption         =   "�H�Υd�L����Ƥ@�ֲ���"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   300
         TabIndex        =   27
         Top             =   2010
         Width           =   2535
      End
      Begin VB.TextBox txtBillMemo 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1800
         MaxLength       =   19
         TabIndex        =   25
         Top             =   3930
         Visible         =   0   'False
         Width           =   6645
      End
      Begin VB.Frame fraOtherBank 
         Caption         =   "�p�X�H�Υd�M�����"
         ForeColor       =   &H00FF0000&
         Height          =   1200
         Left            =   300
         TabIndex        =   20
         Top             =   750
         Width           =   8160
         Begin VB.TextBox txtAuthorizationBatch 
            Height          =   300
            Left            =   1170
            MaxLength       =   8
            TabIndex        =   24
            Text            =   "00001"
            Top             =   285
            Width           =   1575
         End
         Begin VB.OptionButton opyymm 
            Caption         =   "YYMM"
            Height          =   195
            Left            =   6435
            TabIndex        =   7
            Top             =   930
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optmmyy 
            Caption         =   "MMYY"
            Height          =   195
            Left            =   6435
            TabIndex        =   6
            Top             =   660
            Width           =   1095
         End
         Begin VB.ComboBox cboMore 
            Height          =   315
            ItemData        =   "frmUnion.frx":0505
            Left            =   1170
            List            =   "frmUnion.frx":0512
            TabIndex        =   4
            Top             =   675
            Width           =   2460
         End
         Begin VB.CheckBox chkCancel 
            Caption         =   "����"
            Height          =   300
            Left            =   5085
            TabIndex        =   5
            Top             =   660
            Width           =   750
         End
         Begin VB.ComboBox cboSet 
            Height          =   315
            ItemData        =   "frmUnion.frx":0556
            Left            =   1155
            List            =   "frmUnion.frx":0563
            TabIndex        =   2
            Top             =   270
            Width           =   2460
         End
         Begin VB.TextBox txtMerchantName 
            Height          =   300
            Left            =   5085
            MaxLength       =   8
            TabIndex        =   3
            Top             =   255
            Width           =   2460
         End
         Begin VB.Label Label5 
            Alignment       =   1  '�a�k���
            Caption         =   "����X��"
            Height          =   285
            Left            =   225
            TabIndex        =   23
            Top             =   750
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  '�a�k���
            Caption         =   "����X"
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   345
            Width           =   915
         End
         Begin VB.Label Label8 
            Alignment       =   1  '�a�k���
            Caption         =   "�ݥ����N��"
            Height          =   285
            Left            =   3900
            TabIndex        =   21
            Top             =   315
            Width           =   1080
         End
      End
      Begin Gi_Date.GiDate GiDAccept 
         Height          =   285
         Left            =   5145
         TabIndex        =   1
         Top             =   360
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
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
      Begin VB.TextBox txtSpcNO 
         Height          =   300
         Left            =   1455
         TabIndex        =   0
         Top             =   352
         Width           =   2475
      End
      Begin VB.Frame frmData 
         Caption         =   "����ɦ�m"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         HelpContextID   =   2
         Left            =   285
         TabIndex        =   15
         Top             =   2310
         Width           =   8190
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
            Left            =   2625
            TabIndex        =   8
            ToolTipText     =   "�п�J�r�ɤ����|���ɦW�I"
            Top             =   480
            Width           =   4725
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
            Left            =   7620
            TabIndex        =   11
            Top             =   960
            Width           =   375
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
            Left            =   7605
            TabIndex        =   9
            Top             =   495
            Width           =   375
         End
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
            Left            =   2625
            TabIndex        =   10
            Top             =   945
            Width           =   4725
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "����ɦ�m(���|)"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   555
            Width           =   1485
         End
         Begin VB.Label lblErrpath 
            AutoSize        =   -1  'True
            Caption         =   "���D�Ѧ��ɦ�m(���| + �W��)"
            Height          =   195
            Left            =   105
            TabIndex        =   16
            Top             =   990
            Width           =   2460
         End
      End
      Begin VB.Label lblMonthDay 
         Caption         =   "�C��I�ڤ�"
         Height          =   210
         Left            =   6480
         TabIndex        =   28
         Top             =   390
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblBillMessage 
         AutoSize        =   -1  'True
         Caption         =   "�b�洣�C�T��"
         Height          =   195
         Left            =   450
         TabIndex        =   26
         Top             =   4110
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label7 
         Alignment       =   1  '�a�k���
         Caption         =   "�ө��N��"
         Height          =   285
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "�e����"
         Height          =   210
         Left            =   4290
         TabIndex        =   18
         Top             =   390
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�}�l"
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   4320
      Width           =   1875
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   7410
      TabIndex        =   13
      Top             =   4320
      Width           =   1275
   End
End
Attribute VB_Name = "frmUnion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Private cn As New ADODB.Connection
Private rsTmp As New ADODB.Recordset     '' ���X�����Ʈw �����
Private rsDefTmp As New ADODB.Recordset    ''�s��r�ɪ��O���� recordset

Private strPrgName As String  ''�{���W��
Private strPath As String     ''GICMIS1.INI���|
Private strDataName As String  ''�ɮצW��
Private strHeader As String   ''�O�����Y�ɪ���r
Private strChoose As String  ''
Private lngErrCount As Long   '' �O�����~����

Public IsTCB As Boolean
Public IsTCBCard As Boolean
Public strChooseMultiAcc As String   '' �o�ӰѼƥΨөӱ��z�� SO033 �һݪ����O���
Public blnfrmChina  As Boolean ' 2003/10/08   �� SO3272A  �ǤJ�A���ܨ� �� ����H�U�H�Υd�榡��
Dim strBankSQL As String
Dim strBankId As String
Public frmCreditCardType As mCreditCardType     ''  2003/10/28 �o�ӭȥN��ҭn�ϥΪ��H�Υd����
Dim intBatchNumber As Integer  '' �O�� HSBC����X�L���妸
Dim strBatchDate As String     '' �O�� HSBC�X�L�妸�����

Private Sub chkDuteDate_Click()
  On Error Resume Next
    '#4090 �W�[�H�Υd�L���n����Log By Kin 2008/09/19
    If chkDuteDate.Value = 1 Then
        Call MsgBox("��z��ܤĿ�o�ӿﶵ�A�H�Υd�L����ư��F�Q�O�����~�A" & vbCrLf & _
                             "��N�Q�@�ֲ��͡A�Y�O�z������o��@�A�Ш����o�ӤĿ�C", vbInformation, "�L����Ʋ���")
    End If

End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub cmdOK_Click()

 '''If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Call DefineRs   ''�إ߰O���� recordset ���c
 
    Dim sFile As String
    Dim sPath As String
    sFile = txtDataPath.Text
    If frmCreditCardType = lngHSBC Then
        sPath = txtDataPath.Text
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        sFile = sPath & "\" & txtAuthorizationBatch.Text & Right(GiDAccept.GetValue, 6)
    ElseIf frmCreditCardType = lngNewWeb Then   '#5494 �W�[�ŷs����ɦW By Kin 2010/02/03
        sPath = txtDataPath.Text
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        sFile = sPath & txtSpcNO.Text & "_" & GiDAccept.GetValue & "." & txtAuthorizationBatch.Text
    ElseIf frmCreditCardType = lngUnionBatch Then
    
        sPath = txtDataPath.Text
        Dim iSch As Long
        
        iSch = InStrRev(sPath, "\")
        sPath = Mid(sPath, 1, iSch)
        txtDataPath.Text = sPath
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        sFile = sPath & txtMerchantName.Text & ".TXT"
    End If
    
    If OpenFile(sFile, 0) = False Then Exit Sub
    If OpenFile(txtErrPath.Text, 1) = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    objStorePara.uProcText = txtDataPath.Text
    objStorePara.uProcErrText = txtErrPath.Text
    Call ScrToRcd   '#4018 ������ɴN�n�N���Keep�_�� By Kin 2008/07/18
    If Not BeginTran Then
        If Not blnNodata Then
            objStorePara.uUpdate = False
            MsgBox "��b��ƿ�X����!", vbExclamation, "ĵ�i..."
            CloseFS
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
        Else
            objStorePara.uNoData = True
            CloseFS
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
        End If
    End If
    'Call ScrToRcd
    objStorePara.uUpdate = True
    Unload Me
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = vbKeyF2 Then
        cmdOK.Value = cmdOK.Enabled
    End If
  
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
      Me.Caption = objStorePara.BankName & ""
      cboMore.ListIndex = 2
      GiDAccept.Text = Format(Date, "eemmdd")
      Call InitData
      RcdToScr
      If IsTCB = True Then
           Dim rsCompName As New ADODB.Recordset
        '            rsCompName.CursorLocation = adUseClient
        '            rsCompName.Open "SELECT DESCRIPTION FROM " & TableOwnerName & "CD039", cn, adOpenKeyset, adLockReadOnly
           If Not GetRS(rsCompName, "SELECT DESCRIPTION FROM " & TableOwnerName & "CD039", cn) Then Exit Sub
           txtSpcNO.Text = rsCompName(0)
           rsCompName.Close
           Set rsCompName = Nothing
           fraOtherBank.Enabled = False
        ''   txtSpcNO.Locked = True
            
      End If
      
'      If blnfrmChina = False Then
'                 fraOtherBank.Enabled = False
'      End If
      
End Sub
Private Sub RcdToScr()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    Dim interr As Integer
    
    interr = 0
    Set LogFile = Nothing
    interr = 1
    If Dir(strPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
    interr = 2
    Set LogFile = fso.OpenTextFile(strPath & "\" & strPrgName & "Out.log", ForReading)
    interr = 3
    '#5494 �h�W�[Ū���C��I�ڤ� By Kin 2010/02/03
    With LogFile
        If Not .AtEndOfStream Then txtSpcNO = .ReadLine & "":      interr = 4
        If Not .AtEndOfStream Then txtMerchantName = .ReadLine & "":       interr = 5
        If Not .AtEndOfStream Then txtDataPath = .ReadLine & "":        interr = 6 '�����
        If Not .AtEndOfStream Then txtErrPath = .ReadLine & "":        interr = 7 '���D��
        'If Not .AtEndOfStream Then cboMonthDay.Text = .ReadLine & "":    interr = 8
        
    End With
    interr = 71
    If frmCreditCardType = 2 Then
        interr = 72
        If Len(txtDataPath.Text) = 0 Then
            intBatchNumber = 0: interr = 8
            strBatchDate = Format(Date, "eemmdd"): interr = 9
        Else
            interr = 91
            Dim aPath() As String
            Dim varAnayPath
            Dim strAnayPath As String
            aPath = Split(txtDataPath.Text, "\"): interr = 92
            For Each varAnayPath In aPath
                strAnayPath = varAnayPath: interr = 94
            Next
             '' 2004/03/04 ���ѥ[�J�H�U���P�_���A�]���Y�OstrAnayPath�O�Ū��A�|�X�{���~
            If Len(strAnayPath) < 1 Then
                strBatchDate = Format(Date, "eemmdd"): interr = 941
                intBatchNumber = 0: interr = 942
            Else
                aPath = Split(strAnayPath, "."): interr = 95
                strAnayPath = aPath(0): interr = 96
                If Right(strAnayPath, 6) <> Format(Date, "eemmdd") Then
                    strBatchDate = Format(Date, "eemmdd"): interr = 97
                    intBatchNumber = 0: interr = 98
                Else
                    strBatchDate = Right(strAnayPath, 6): interr = 99
                    intBatchNumber = CInt(Mid(strAnayPath, 2, Len(strAnayPath) - 7)): interr = 991
                End If
                interr = 10
            End If
        End If
    End If
    interr = 11
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr" & "   errcode :" & interr)
End Sub

Private Sub InitData()
  On Error GoTo ChkErr
'  Dim o As Object
'  Set o = CreateObject("CreditCardOut.clsInterface")
    With objStorePara
        Set cn = .Connection
        strBankId = .BankId
        '#3531 �h�W�[�Ȧ����(strBankId����g�o�Ӧ�) By Kin 2007/11/02
        strBankSQL = .BankStr
        'strBankName = .BankName
        'strCorpId = .CorpID
        strChoose = .ChooseStr
        strPath = .ErrPath
        txtSpcNO.Text = .uSpcNO
    End With
    If frmCreditCardType <> lngNewWeb Then
        If frmCreditCardType <> lngHSBC Then lblDatapath.Caption = "����ɦ�m(���| + �W��)"
    End If
    
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "InitData")
End Sub

Public Property Get PrgName() As String
    PrgName = strPrgName
End Property

Public Property Let PrgName(ByVal vData As String)
    strPrgName = vData
End Property


Private Sub cmdDataPath_Click()
  On Error GoTo ChkErr
    '#5494 �W�[�ŷs��ު��P�_ By Kin 2010/02/03
    If (frmCreditCardType = lngHSBC) Or (frmCreditCardType = lngNewWeb) Or (frmCreditCardType = lngUnionBatch) Then
        txtDataPath.Text = FolderDialog("�}�Ҹ��|")
'        If frmCreditCardType = lngUnionBatch Then
'            If txtMerchantName.Text <> "" Then
'                If Right(txtDataPath.Text, 1) = "\" Then
'                    txtDataPath.Text = txtDataPath & txtMerchantName.Text & ".TXT"
'                Else
'                    txtDataPath.Text = txtDataPath & "\" & txtMerchantName.Text & ".TXT"
'                End If
'            End If
'        End If
    Else
        With cnd1
            .FileName = txtErrPath.Text
            .Filter = "��r��|*.txt"
            .InitDir = strPath
            .Action = 1
            txtDataPath = .FileName
        End With
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdDataPath_Click")
End Sub

Private Sub cmdErrPath_Click()
    On Error GoTo ChkErr
    With cnd1
        .FileName = txtErrPath.Text
        .Filter = "��r��|*.txt"
        .InitDir = strPath
        .Action = 1
        txtErrPath = .FileName
    End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdErrPath_Click")
End Sub

Private Sub DefineRs()
    On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        Select Case frmCreditCardType
        Case 0
            If IsTCB = False Then
            ''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
                .Fields.Append ("F110"), adBSTR, 10, adFldIsNullable
                .Fields.Append ("F1118"), adBSTR, 8, adFldIsNullable
                .Fields.Append ("F1937"), adBSTR, 19, adFldIsNullable
                .Fields.Append ("F3845"), adBSTR, 8, adFldIsNullable
                .Fields.Append ("F4653"), adBSTR, 8, adFldIsNullable
                
                .Fields.Append ("F5455"), adBSTR, 2, adFldIsNullable
                .Fields.Append ("F5661"), adBSTR, 6, adFldIsNullable
                .Fields.Append ("F6265"), adBSTR, 4, adFldIsNullable
                .Fields.Append ("F6668"), adBSTR, 3, adFldIsNullable
                .Fields.Append ("F6984"), adBSTR, 16, adFldIsNullable
                
                .Fields.Append ("F8590"), adBSTR, 6, adFldIsNullable
                .Fields.Append ("F9196"), adBSTR, 6, adFldIsNullable
                .Fields.Append ("F97112"), adBSTR, 16, adFldIsNullable
                .Fields.Append ("F113118"), adBSTR, 6, adFldIsNullable
                .Fields.Append ("F119"), adBSTR, 1, adFldIsNullable
                
                .Fields.Append ("F120"), adBSTR, 1, adFldIsNullable
                    
            Else   '' // �H�U�����ذӻȪ��榡
                If IsTCBCard = True Then
                    .Fields.Append ("F06"), adBSTR, 6, adFldIsNullable   '' �дڤ��
                    .Fields.Append ("F10"), adBSTR, 4, adFldIsNullable   '' �Ǹ�
                    .Fields.Append ("F26"), adBSTR, 16, adFldIsNullable  '' �d��
                    .Fields.Append ("F29"), adBSTR, 3, adFldIsNullable   '' CVC2
                    .Fields.Append ("F33"), adBSTR, 4, adFldIsNullable   '' ���Ħ~��
                    .Fields.Append ("F44"), adBSTR, 11, adFldIsNullable  ''�q��s��
                       
                    .Fields.Append ("F50"), adBSTR, 6, adFldIsNullable   ''������
                    .Fields.Append ("F56"), adBSTR, 6, adFldIsNullable   '' ���v�X
                    .Fields.Append ("F58"), adBSTR, 2, adFldIsNullable   '' ����
                    .Fields.Append ("F65"), adBSTR, 7, adFldIsNullable   '' �������B
                    .Fields.Append ("F72"), adBSTR, 7, adFldIsNullable    '' �`���B
                    .Fields.Append ("F79"), adBSTR, 7, adFldIsNullable    '' �Y ��
                    
                    .Fields.Append ("F86"), adBSTR, 7, adFldIsNullable  ''  ����
                    .Fields.Append ("F91"), adBSTR, 5, adFldIsNullable   '' �S���N��
                    .Fields.Append ("F116"), adBSTR, 25, adFldIsNullable  '' �S���W��
                    .Fields.Append ("F136"), adBSTR, 20, adFldIsNullable  '' �S���W��
                    .Fields.Append ("F138"), adBSTR, 2, adFldIsNullable   ''  ������O
                    
                    .Fields.Append ("F180"), adBSTR, 62, adFldIsNullable   ''  �ť�
                Else
                    .Fields.Append ("F15"), adBSTR, 16, adFldIsNullable   '' �q��s��
                    .Fields.Append ("F16"), adBSTR, 15, adFldIsNullable   '' �P�f�s��
                    .Fields.Append ("F31"), adBSTR, 19, adFldIsNullable  '' �d��
                    .Fields.Append ("F50"), adBSTR, 3, adFldIsNullable   '' CVC2
                    .Fields.Append ("F53"), adBSTR, 4, adFldIsNullable   '' ���Ħ~��(YYMM)
                    .Fields.Append ("F57"), adBSTR, 10, adFldIsNullable  ''������B
                    .Fields.Append ("F67"), adBSTR, 2, adFldIsNullable  ''���]��
                    .Fields.Append ("F69"), adBSTR, 2, adFldIsNullable   ''����
                    .Fields.Append ("F71"), adBSTR, 12, adFldIsNullable   '' �Y��
                    .Fields.Append ("F83"), adBSTR, 12, adFldIsNullable  ''����
                End If
            End If
        Case 1
           ''  2003/10/08   �p�G�O����ӻȪ��飼�H�U�o�@�q
    
            ''''�S���N���G�̷~�Ȧ����P�N�� , �|���T�w
            ''''�d���G�p�X�H�Υd�ɥd���u��11��: ����4�� '4000',�A��11��d��,�A��1��'0',�զ�16�� ; ��L�d��16��
            ''''����X : ��l ��0 ��+�ť� , �g���v�B�z��|�ܬ� ��1��+�ť� , �j�����v(�G�����v)��|�ܬ� ��2��+�ť�
            ''''�y����:  �ӵ����Ӧb�ɮפ��ĴX�� , ���藍�୫�_
            ''''�^���X : ��1+�ť�+�ťա��Ρ� 5+�ť�+�ťա��ݦ��\���,��������^���T�����w��APPROVED ; �������~���^���X���ݱ��v���ѥ��
            ''''�^���T�� : ��APPROVED���N���\ ; ��REFERRAL���N��CALL BANK,��L�h�N����,�̱`����TRANSACTION ERRO���N���ƿ��~ ; ��INVALID CAPTURE���N���L���� ; ��DECLINE���N���L����,�Բӭ�]�наѷӡ��妸�^���X���@��

            ''  �H�U�Хܬ��@"**"  ��ܥ��ݥѵ{����J�A��L�h�O�Ȧ�^���X
            
            .Fields.Append ("F113"), adBSTR, 13, adFldIsNullable   ''  ** �S���N��(�����p�X:�ݥ����N��)
            .Fields.Append ("F1429"), adBSTR, 16, adFldIsNullable  ''  ** �d��
            .Fields.Append ("F3039"), adBSTR, 10, adFldIsNullable  ''  ** ���B
            .Fields.Append ("F4043"), adBSTR, 4, adFldIsNullable   ''   ** ���Ĥ�~ (MMYY) (�褸�~)
            .Fields.Append ("F4445"), adBSTR, 2, adFldIsNullable   '''   ** ����X   (0+�ť�)
            .Fields.Append ("F4659"), adBSTR, 14, adFldIsNullable  '''   ** �y����(���藍�i����)  (������0)
             '' �H�U���Ȧ�^�ХΪ�
            .Fields.Append ("F6065"), adBSTR, 6, adFldIsNullable  ''  APPROVED CODE
            .Fields.Append ("F6671"), adBSTR, 6, adFldIsNullable   ''   APPROVED DATE (YYMMDD) (�褸)
            .Fields.Append ("F7274"), adBSTR, 3, adFldIsNullable   ''  �^���X
            .Fields.Append ("F7590"), adBSTR, 16, adFldIsNullable  '''  �^���T��
            
            .Fields.Append ("F9196"), adBSTR, 6, adFldIsNullable   ''  ���v���(APPROVED DATE)
            .Fields.Append ("F97102"), adBSTR, 6, adFldIsNullable   ''  ���v�ɶ�(��/��/��)
        Case 2
                ''  2003/11/20  �H�U���׻Ȧ�H�Υd�榡
            .Fields.Append ("F116"), adBSTR, 16, adFldIsNullable   ''  ** �d��  �����A�̨t�γW�w�ɨ�16�Xspace
            .Fields.Append ("F1717"), adBSTR, 1, adFldIsNullable  ''  ** �ť�
            .Fields.Append ("F1821"), adBSTR, 4, adFldIsNullable  ''  ** ���Ĵ���  MBF,MMYY(�褸�~)
            .Fields.Append ("F2222"), adBSTR, 1, adFldIsNullable   ''   ** " " or "-"
            .Fields.Append ("F2334"), adBSTR, 12, adFldIsNullable   '''   ���B  ������'0'���X������
            .Fields.Append ("F3535"), adBSTR, 1, adFldIsNullable  '''  �ť�
        
            .Fields.Append ("F3641"), adBSTR, 6, adFldIsNullable  ''   �ֲa���X   ��J'999999'or approvaled code
            .Fields.Append ("F4242"), adBSTR, 1, adFldIsNullable   ''   �ť�
            .Fields.Append ("F4344"), adBSTR, 2, adFldIsNullable   ''  �^�_�X   ��J'99'or '0'(if this is a refund)
            .Fields.Append ("F4545"), adBSTR, 1, adFldIsNullable  '''  �ť�
            .Fields.Append ("F4685"), adBSTR, 40, adFldIsNullable   ''  �ө��ۦ�B�� (so043.para24=1�h���C��渹
        Case 3
            .Fields.Append ("SPCNO"), adBSTR, 5, adFldIsNullable
            .Fields.Append ("ACCOUNTID"), adBSTR, 19, adFldIsNullable
            .Fields.Append ("STOPYM"), adBSTR, 4, adFldIsNullable
            .Fields.Append ("SHOULDAMT"), adBSTR, 13, adFldIsNullable
            .Fields.Append ("BILLNO"), adBSTR, 15, adFldIsNullable
            .Fields.Append ("TRADETYPE"), adBSTR, 1, adFldIsNullable
            .Fields.Append ("CHINESE"), adBSTR, 40, adFldIsNullable
        '#5494 �W�[�ŷs��ޮ榡 By Kin 2010/01/27
        Case 5
            .Fields.Append "RcdType", adBSTR, 10, adFldIsNullable   '�O�����O,�T�w��D
            .Fields.Append "OrderNo", adBSTR, 30, adFldIsNullable   '�ө��q��渹
            .Fields.Append "OrdItem", adBSTR, 9, adFldIsNullable    '�q��s�� (����Ǹ�)
            .Fields.Append "Period", adBSTR, 2, adFldIsNullable     '��������
            .Fields.Append "Status", adBSTR, 1, adFldIsNullable     '������A
            .Fields.Append "Amt", adBSTR, 13, adFldIsNullable       '�q����B(�t2��p���I)
            .Fields.Append "CardNo", adBSTR, 16, adFldIsNullable    '�H�Υd�d��
            .Fields.Append "CVC2", adBSTR, 3, adFldIsNullable       'CVC2
            .Fields.Append "ValidDate", adBSTR, 6, adFldIsNullable  '���Ħ~��
            .Fields.Append "AuthNo", adBSTR, 6, adFldIsNullable     '���v�X
            .Fields.Append "Prc", adBSTR, 5, adFldIsNullable        '�D�^�нX
            .Fields.Append "Src", adBSTR, 5, adFldIsNullable        '���^�нX
            .Fields.Append "RetBank", adBSTR, 2, adFldIsNullable    '�Ȧ�^�нX
            .Fields.Append "BatNum", adBSTR, 9, adFldIsNullable     '�妸���X
            .Fields.Append "AutoFlag", adBSTR, 1, adFldIsNullable   '�۰ʽдںX��
            .Fields.Append "hold", adBSTR, 91, adFldIsNullable      '�O�d���
            
        '#5609 �W�[�p�X�H�Υd�妸 By Kin 2010/06/03
        Case 6
            Call DefRs_6(rsDefTmp)
        End Select
        .Open
    End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefineRs")
End Sub
Private Sub DefRs_6(ByRef rsDef As Recordset)
  On Error GoTo ChkErr
    With rsDef
        .Fields.Append "SPCNO", adBSTR, 10, adFldIsNullable         '�ө��N�X
        .Fields.Append "MERCHANTNAME", adBSTR, 8, adFldIsNullable   '�ݥ����N��
        .Fields.Append "ACCOUNTNO", adBSTR, 16, adFldIsNullable     '�d��
        .Fields.Append "ACCOUNTNOEX", adBSTR, 3, adFldIsNullable    '�d�������X
        .Fields.Append "STOPYM", adBSTR, 4, adFldIsNullable         '���Ĵ���
        .Fields.Append "AMT", adBSTR, 10, adFldIsNullable           '������B
        .Fields.Append "AUTHNO", adBSTR, 8, adFldIsNullable         '���v�X
        .Fields.Append "SETNO", adBSTR, 2, adFldIsNullable          '����X
        .Fields.Append "ACCEPTDATE", adBSTR, 6, adFldIsNullable     '�e����
        .Fields.Append "BILLNO", adBSTR, 16, adFldIsNullable        'USER DEFINE
        .Fields.Append "ADDITION", adBSTR, 30, adFldIsNullable      'Additional field
        .Fields.Append "ACCINFO", adBSTR, 40, adFldIsNullable       '�d�H��T
    End With
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "DefRs_6")
End Sub
Private Function BeginTran() As Boolean
  On Error GoTo ChkErr
    Dim strsql As String
    Dim strOldBillNo As String
    Dim blnFirst As Boolean
    
    Dim lngErrCount As Long
    Dim lngCount As Long
    Dim lngTime As Long
    Dim lngIndex As Long
    
    Dim strSequenceNumber As String   '' ���o�����즹�渹������Ǹ�
    Dim intPara24 As Integer  ''  �O���O�_�ϥδC��h�b��B�z
    Dim strMediaBillNOField As String   '' �O���ثe�� BillNO �άO  MediaBillNO
    Dim rsMedioBillnoNull As New Recordset   '' �o�ӰѼƥΨөӱ��z�� SO033�� mediobillno �Onull �����
    Dim strMedioBillnoNull As String    '' �o�ӰѼƥΨөӱ��z�� SO033�� mediobillno �Onull �����
    Dim rsCustIDErr As New ADODB.Recordset  ''2003/08/17
    Dim lngTotalHSBC As Long: lngTotalHSBC = 0
    Dim lngToatalCountHSBC As Long: lngToatalCountHSBC = 0
    Dim lngTotalHSBCN As Long:   lngTotalHSBCN = 0 ''�O�����B���t���`��
    Dim lngTotalCountHSBCN As Long:   lngTotalCountHSBCN = 0  ''�O�����B���t������
    Dim begintrnsErrCode As Double
    Dim strStopYmSQL As String
    Dim lngShouldAmt As Long
    Dim rsSO106 As New ADODB.Recordset
    Dim rsCD013 As New ADODB.Recordset
    Dim rsUpdUCCode As New ADODB.Recordset
    Dim strUCCode As String
    Dim strUCName As String
    Dim strUpdUCCode As String
    Dim blnUpdUCCode As Boolean
    Dim strUCCodeWhere As String
    Dim strStopYmWhere As String
    begintrnsErrCode = 0
    lngShouldAmt = 0
    
    '********************************************************************************************************************************************************
    '#3417 �q�l�ɶץX��,�n��J������](RefNo=4) By Kin 2007/12/05
    If Not GetRS(rsCD013, "Select * From " & TableOwnerName & "CD013 Where RefNo=4 And StopFlag<>1 Order By CodeNo Desc", cn, adUseClient, adOpenKeyset) Then Exit Function
    If rsCD013.EOF Then
        blnUpdUCCode = False
    Else
        blnUpdUCCode = True
        strUCCode = rsCD013("CodeNo") & ""
        strUCName = rsCD013("Description") & ""
    End If
    '*********************************************************************************************************************************************************
    strStopYmWhere = " Max(Decode(Length(C.StopYm),5," & _
                " Substr(C.StopYm,2,4) || '0' || Substr(C.StopYm,1,1), " & _
                " Substr(C.StopYm,3,4) || Substr(C.StopYm,1,2))) CardExpDate"

    '#3728 �j���H�h�C�^�X�b By Kin 2008/03/25
    'intPara24 = CInt(RPxx(cn.Execute("SELECT Para24 FROM " & TableOwnerName & "SO043 Where Rownum=1").GetString))
    intPara24 = 1
    If intPara24 = 0 Then
        strMediaBillNOField = "BillNO"
    Else
        strMediaBillNOField = "MediaBillNO"
    End If
     BeginTran = False
     lngTime = Timer
     strsql = ""
        
 '' �o�����   "TCBbudget "  �ΥH���o��������
    strsql = ""
 
 ''2003/08/18  �N�䤤��CUSTID ���� �A�ѩ��U���s���o
    '#5267 �쥻��SO002A�{�b���SO106 By Kin 2010/04/20
    If IsTCB = False Then    '' �p�G�O���ذӻȪ��榡�ASQL �y�k�٭n�N�䤤�� CitemCode �o�Ӷ��ظs�եX��
    ''  ���׭n�̥d���@�Ƨ� �A�ҥH�s�W�H�����{���X
        Dim strOrderbyAccountNO As String
        strOrderbyAccountNO = strMediaBillNOField
        If frmCreditCardType = lngHSBC Then
             strOrderbyAccountNO = "AccountNO  DESC  "
        End If
    '' errcode**********
    '' 20050203 ����And A.ShouldAmt > 0�����󭭨�
        begintrnsErrCode = 1
        '#5055 ���M�S������,���O���K�N�d�x�w���L�o�� By Kin 2009/04/30
        '#5218 �d�x�w���n��Nvl(CD013.REFNO,0) By Kin 2009/08/05
        '#5564 �ѦҸ�7,8�PCD013.PAYOK=1���O�N��w�� By Kin 2010/05/17
        '#5683 �W�[�D��RealStopDate �P PayKind By Kin 2010/08/06
        strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
            " A." & strMediaBillNOField & ", sum(A.ShouldAmt)  ShouldAmt," & _
            " A.AccountNO,B.CUSTNAME,B.CUSTID," & _
            " SUM(A.TCBbudget) TCBbudget , MAX(A.PTCODE)  PtCode," & _
            " MAX(A.ServiceType) ServiceType,  " & _
            strMinPayKind & " , " & strMinRealStopDate & _
            " FROM " & TableOwnerName & "SO033 A," & TableOwnerName & "SO001 B ," & TableOwnerName & "SO002 D " & _
            "," & TableOwnerName & "CD013 " & _
            "Where " & strBankId & IIf(Len(strBankId) > 0, " AND ", " ") & _
            "A.CustID  = D.CUSTID AND " & _
            "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
            "A.UCCode Is Not Null  AND " & _
            "A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) AND " & _
            "NVL(CD013.PAYOK,0)= 0 " & _
            IIf(Len(strChoose) = 0, "", " And ") & strChoose & " AND " & _
            "A.SERVICETYPE = D.SERVICETYPE   " & _
            "Group by A." & strMediaBillNOField & " ,  A.AccountNO  , B.CUSTNAME,B.CUSTID  ORDER BY A." & strOrderbyAccountNO
                
        strsql = "SELECT A." & strMediaBillNOField & " BILLNO ,ShouldAmt," & _
            " A.AccountNO,A.CUSTNAME,A.CUSTID,Min(A.RealStopDate) RealStopDate," & _
            " Min(A.PayKind) PayKind , " & _
            " A.TCBbudget , A.PTCODE,A.ServiceType,C.CVC2, " & strStopYmWhere & _
            " FROM (" & strsql & ") A, " & TableOwnerName & "SO106 C " & _
            " WHERE A.ACCOUNTNO=C.ACCOUNTID AND A.CUSTID=C.CUSTID " & _
            " AND C.STOPFLAG<>1 AND C.SnactionDate IS NOT NULL " & _
            " GROUP BY A." & strMediaBillNOField & " ,ShouldAmt,A.AccountNO," & _
              " A.CUSTNAME,A.CUSTID,A.TCBbudget,A.PTCODE,A.ServiceType,C.CVC2 " & _
            " ORDER BY A." & strOrderbyAccountNO
            
            
            
    Else

        '#5564 �ѦҸ�7,8�PCD013.PAYOK=1���O�N��w�� By Kin 2010/05/17
        strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                   " A." & strMediaBillNOField & ", sum(A.ShouldAmt)  ShouldAmt," & _
                   " A.AccountNO,B.CUSTNAME,B.CUSTID," & _
                   " A.CitemCode," & strMinPayKind & " , " & strMinRealStopDate & " , " & _
                   " MAX(A.PTCODE)  PtCode,MAX(A.ServiceType) ServiceType," & _
                   " MAX(A.TCBbudget) TCBbudget  " & _
                   " FROM " & TableOwnerName & "SO033 A," & TableOwnerName & "SO001 B ," & TableOwnerName & "SO002 D " & _
                   "," & TableOwnerName & "CD013 " & _
                   " Where " & strBankId & IIf(Len(strBankId) > 0, " AND ", " ") & _
                   " A.CustID  = D.CUSTID AND " & _
                   " A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                   " A.UCCode Is Not Null  And " & _
                   " A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) AND  " & _
                   " NVL(CD013.PAYOK,0) = 0 " & _
                   IIf(Len(strChoose) = 0, "", " And ") & strChoose & " AND " & _
                   " A.SERVICETYPE = D.SERVICETYPE " & _
                   " Group by A." & strMediaBillNOField & " ,  A.AccountNO ,  B.CUSTNAME , A.CItemCode,B.CUSTID  " & _
                   " ORDER BY A." & strMediaBillNOField
                   
        strsql = "SELECT A." & strMediaBillNOField & " BILLNO ,A.ShouldAmt," & _
                " A.AccountNO,A.CUSTNAME,Min(A.PAYKIND) PAYKIND,MIN(REALSTOPDATE) REALSTOPDATE ," & _
                " A.CUSTID,A.CitemCode,A.PtCode,A.ServiceType,TCBbudget,C.CVC2, " & strStopYmWhere & _
                " From (" & strsql & ") A," & TableOwnerName & "SO106 C " & _
                " WHERE A.ACCOUNTNO=C.ACCOUNTID AND A.CUSTID=C.CUSTID " & _
                " AND C.STOPFLAG<>1 AND C.SnactionDate IS NOT NULL " & _
                " GROUP BY A." & strMediaBillNOField & " ,ShouldAmt,A.AccountNO," & _
                " A.CUSTNAME,A.CUSTID,A.CitemCode,A.PtCode,A.ServiceType,A.TCBbudget,C.CVC2 " & _
                " ORDER BY A." & strMediaBillNOField
        
    End If
    
    '*********************************************************************************************
    '#3417 ��sUCCode�PUCName����Where ���� By Kin 2007/12/05
    '#5267 C.�쥻�OSO002A,���SO106
    strUCCodeWhere = strBankId & IIf(Len(strBankId) > 0, " AND ", " ") & _
                      "A.CustID  = C.CUSTID AND " & _
                      "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                      "A.UCCode Is Not Null  " & _
                      IIf(Len(strChoose) = 0, "", " And ") & strChoose & " AND " & _
                      "A.AccountNO = C.AccountID AND A.CustID = D.CustID  AND " & _
                      "A.SERVICETYPE = D.SERVICETYPE " & _
                      " AND C.StopFlag<> 1 AND C.SnactionDate IS NOT NULL "
    '*********************************************************************************************

    ''  2003/08/14 �令 �A�N mediabillno   ��ƨ��X �o�سo�@���h���@�� BillNO �M���J��mediabillno��null �Ȫ��ɭԡA��J�C��渹
    '' 2003/08/17  ���ݦA�@�@���A��¨��XmediaBillno�Onull���ȡA�åB�N���J�s��
    '' errcode**********
    begintrnsErrCode = 2
    If intPara24 = 1 Then
        ''�o�@�q�������e�z���D SQL (strsql)  ���� group  �A�[�W MediaBillno
        '' errcode**********
        '' 20050203 ����And A.ShouldAmt > 0�����󭭨�
        '#5055 �S����,�����K�@�_��,�L�o�d�x�w������� By Kin 2009/04/30
        '#5218 �d�x�w���n��Nvl By Kin 2009/08/05
        '#5564 �ѦҸ�7,8�PCD013.PAYOK=1���O�N��w�� By Kin 2010/05/17
        '#5683 �W�[�D��RealStopDate �P PayKind By Kin 2010/08/06
        begintrnsErrCode = 2
        If IsTCB = False Then    '
            strMedioBillnoNull = "SELECT A.BillNO, A.MediaBillNO,A.AccountNO," & _
                        strMinPayKind & " , " & strMinRealStopDate & " , " & strStopYmWhere & _
                        " FROM " & TableOwnerName & "SO033 A," & TableOwnerName & "SO001 B ," & TableOwnerName & "SO002 D," & TableOwnerName & "SO106 C  " & _
                        "," & TableOwnerName & "CD013 " & _
                        " Where " & strBankId & IIf(Len(strBankId) > 0, " AND ", " ") & _
                        " A.CustID  = D.CUSTID AND " & _
                        " A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                        " A.UCCode Is Not Null  And " & _
                        " A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) AND  " & _
                        " NVL(CD013.PAYOK,0) = 0 " & _
                        IIf(Len(strChoose) = 0, "", " And ") & strChoose & " AND " & _
                        " A.SERVICETYPE =D.SERVICETYPE  AND   " & _
                        " A.AccountNO = C.AccountID AND A.CustID = C.CustID  AND A.MediaBillNO IS NULL " & _
                        " AND C.StopFlag <> 1 AND C.SnactionDate Is Not Null " & _
                        " Group by A.BillNo,A.MediaBillNo,A.AccountNo "
        Else
            strMedioBillnoNull = "SELECT A.BillNO,A.MediaBillNO,A.AccountNO," & _
                        strMinPayKind & " , " & strMinRealStopDate & " , " & strStopYmWhere & _
                       " FROM " & TableOwnerName & "SO033 A," & TableOwnerName & "SO001 B ," & TableOwnerName & "SO002 D," & TableOwnerName & "SO106 C  " & _
                       "," & TableOwnerName & "CD013 " & _
                       " Where " & strBankId & IIf(Len(strBankId) > 0, " AND ", " ") & _
                       " A.CustID  = D.CUSTID AND " & _
                       " A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                       " A.UCCode Is Not Null  And " & _
                       " A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT INT (3,7) AND  " & _
                       " NVL(CD013.PAYOK,0) = 0 " & _
                       IIf(Len(strChoose) = 0, "", " And ") & strChoose & " AND " & _
                       " A.SERVICETYPE =D.SERVICETYPE  AND   " & _
                       " A.AccountNO = C.AccountID AND A.CustID = C.CustID  AND A.MediaBillNO IS NULL " & _
                       " AND C.StopFlag <> 1 AND C.SnactionDate Is Not Null " & _
                       " Group by A.BillNo,A.MediaBillNo,A.AccountNo "
        End If
        '' errcode**********
        '' 20050203 ����And A.ShouldAmt > 0�����󭭨�
        
        begintrnsErrCode = 3
        strChooseMultiAcc = "  A.CancelFlag = 0 AND " & _
                            "A.UCCode Is Not Null   " & _
                             IIf(Len(strChooseMultiAcc) = 0, "", " And ") & strChooseMultiAcc
        '' errcode**********
        begintrnsErrCode = 4
        If GetRS(rsMedioBillnoNull, strMedioBillnoNull, cn) Then
            Do While Not rsMedioBillnoNull.EOF
                
                If IsNull(rsMedioBillnoNull("MediaBillNo")) Then
                     strSequenceNumber = GetInvoiceNo2("SO033_MediaBillNo", cn)
                     cn.Execute ( _
                                    "UPDATE " & TableOwnerName & "SO033 A SET " & _
                                    "A.MediaBillNo  ='" & strSequenceNumber & "'  WHERE   " & _
                                    "A.BillNo = '" & rsMedioBillnoNull("BillNO") & "'   AND  " & strChooseMultiAcc & " AND " & _
                                    "A.AccountNO ='" & rsMedioBillnoNull("AccountNO") & "' ")
                   
                End If
                rsMedioBillnoNull.MoveNext
             Loop
        End If
    '' errcode**********
        begintrnsErrCode = 5
        rsMedioBillnoNull.Close
        Set rsMedioBillnoNull = Nothing
    End If
    


    ''********************************************************************************************
  '' errcode**********
    begintrnsErrCode = 6
    Set rsTmp = cn.Execute(strsql)
    lngIndex = 1
    blnFirst = True
  '' errcode**********
    begintrnsErrCode = 7
    ''�o�@�q���o�u�� SO033 ���z����  ........
    '' 20050203 ����And A.ShouldAmt > 0�����󭭨�
    If intPara24 = 0 Then
        strChooseMultiAcc = "  A.CancelFlag = 0 AND " & _
                            "A.UCCode Is Not Null   " & _
                             IIf(Len(strChooseMultiAcc) = 0, "", " And ") & strChooseMultiAcc
    End If
    '' errcode**********
    begintrnsErrCode = 8
    If rsTmp.BOF Or rsTmp.EOF Then
        MsgBox "�L��ơI�Э��s�ާ@�I", vbInformation, "�T���I"
        blnNodata = True
        Unload Me
    Else
        blnNodata = False
        Dim strSN As String  '' ���w�ǦC��
        Dim lngtotal As Long '' �֭p���Y�� Total Amount ����`�X����B
        lngCount = 0
        lngtotal = 0
        Dim lngDoubleMediaNO As Long '' �o�ӭȰO�����ХX�檺��l���
        Dim rsDoubleMediaNO As New ADODB.Recordset    ' �o��   rs �ΥH check  MediaBillNO�O���O����
        lngDoubleMediaNO = 0
 '' errcode**********
        begintrnsErrCode = 9
           ''�}�l�v����������
           ''**********************************************************************************
        While Not rsTmp.EOF
            Dim strCardExpDate As String
            Dim iCED As Integer
            Dim dteCED As Date
            Dim strErrMessage As String
            'rsUpdUCCode.Filter = "BillNo='" & rsTmp("BillNo") & "' And Account"
       '' �o�@�q�ɤW��L���z�����ȡA�]���|���h�� SO033  �P�s��  2003/07/21
                
      '' 2003/08/17  �ѩ�Dsql�q���y�k�ح��u��Group �Ȥ᪺�b���H�γ�ڽs���A
      '' �]�� ���s�A���@���Ƚs�H�ά�������T
'' errcode**********
            begintrnsErrCode = 10
            
    ' Jacky 94/10/05
            Call GetRS(rsCustIDErr, "SELECT  SO001.CUSTNAME , SO033.CUSTID FROM " & TableOwnerName & "SO001," & TableOwnerName & "SO033 " & _
                        "WHERE SO033.CUSTID = SO001.CUSTID And SO033.CompCode = SO001.CompCode And SO033." & strMediaBillNOField & "='" & rsTmp("BillNO") & "' ", cn)
            '#5683 �p�G�����I���p��e�����󤣭n�X�b By Kin 2010/08/06
            If Not IsPayKindOK(rsTmp, GiDAccept.GetValue & "", giAll, 1) Then
                lngErrCount = lngErrCount + 1
                strErrMessage = "�����I��餣���T�G�渹�@" & rsTmp("BillNo") & " �Ƚs�G" & rsCustIDErr("CUSTID") & _
                            " �Ȥ�m�W�@" & rsCustIDErr("custname") & _
                            " �����I���G" & rsTmp("RealStopDate") & " ú�I���O�G�{�I�� "
                file(1).WriteLine strErrMessage
                GoTo NextLoop
                            
            End If
'' errcode**********
            begintrnsErrCode = 11
            If IsNull(rsTmp("AccountNO")) Then
                lngErrCount = lngErrCount + 1
                strErrMessage = "�H�Υd�d���ť� : �渹�@" & rsTmp("BillNO") & " �Ȥ�m�W " & rsCustIDErr("custname")
                file(1).WriteLine strErrMessage
                GoTo NextLoop
            End If
'' errcode**********
            begintrnsErrCode = 12
            '#3531 ���SO106��StopYM(�̤p��) By Kin 2007/11/01
            '#4018 ���SO106���̤j��(�W�[Desc) By Kin 2008/07/18
            '#5494 �ŷs��ޤ]��SO106 By Kin 2010/02/02
            If frmCreditCardType = 3 Or frmCreditCardType = 5 Then
                strStopYmSQL = "Select Decode(Length(A.StopYm),5,Substr(A.StopYm,2,4)||'0'||Substr(A.StopYm,1,1)," & _
                                             "Substr(A.StopYm,3,4)||Substr(A.StopYm,1,2)) As CardExpDate " & _
                                        " From " & TableOwnerName & "SO106 A " & _
                                        "Where A.AccountID='" & rsTmp("AccountNo") & "' " & _
                                        "And A.CustId=" & rsCustIDErr("Custid") & " And A.StopFlag=0 " & _
                                        IIf(strBankId <> "", " And " & strBankId, "") & _
                                        " Order By CardExpDate Desc"
                '#5267 �w�g���SO106�F,�ҥH���Φb�����@�� By Kin 2010/04/20
'                If Not GetRS(rsSO106, strStopYmSQL, cn) Then Exit Function
'                If rsSO106.EOF Then
'                    strCardExpDate = "NULL"
'                Else
'                    strCardExpDate = Left(rsSO106("CardExpDate"), 4) & "/" & Right(rsSO106("CardExpDate"), 2)
'                End If
                strCardExpDate = Empty
                If rsTmp("CardExpDate") & "" <> "" Then
                    strCardExpDate = Left(rsTmp("CardExpDate"), 4) & "/" & Right(rsTmp("CardExpDate"), 2)
                End If
            Else
                strCardExpDate = Empty
                strCardExpDate = rsTmp("CardExpDate") & ""
                'strCardExpDate =  GetNullString(rsTmp("CardExpDate"), giStringV, giAccessDb)
                If UCase(strCardExpDate) = "NULL" Then strCardExpDate = ""
                If strCardExpDate <> "" Then
                    strCardExpDate = Right(strCardExpDate, 2) & Left(strCardExpDate, 4)
                End If
            End If
            If UCase(strCardExpDate) = "NULL" Or strCardExpDate = Empty Then
                lngErrCount = lngErrCount + 1
                strErrMessage = "�H�Υd��Ƥ����T : �渹�@" & rsTmp("BillNO") & " �Ȥ�m�W " & rsCustIDErr("custname")
                file(1).WriteLine strErrMessage
                GoTo NextLoop
            End If
            
'' errcode**********
            begintrnsErrCode = 12.5
            '#3531 �p�G�O�s�榡�ϥ�SO106.StopYm,���O�h�ϥέ�Ӫ���� By Kin 2007/11/01
            If frmCreditCardType <> 3 And frmCreditCardType <> 5 Then
                strCardExpDate = rsTmp("CardExpDate")
                begintrnsErrCode = 12.51
                iCED = IIf(Len(strCardExpDate) = 5, 1, 2)
                strCardExpDate = Left(strCardExpDate, 4) & "/" & Right(strCardExpDate, iCED)
                'strCardExpDate = Right(strCardExpDate, 4) & "/" & Mid(strCardExpDate, 1, iCED)
            End If
            dteCED = CDate(strCardExpDate)
            begintrnsErrCode = 12.52
            If dteCED < CDate(Format(GiDAccept.Text, "YYYY/MM")) Then
                lngErrCount = lngErrCount + 1
                begintrnsErrCode = 12.53
                On Error Resume Next
                If Not rsCustIDErr.EOF Then
                    strErrMessage = "�H�Υd�L�� : �渹�@" & rsTmp.Fields("Billno") & " �Ȥ�m�W�@" & rsCustIDErr("custname") & " �H�Υd�� " & rsTmp.Fields("AccountNO") & " �����@" & _
                                    IIf(frmCreditCardType <> 3, rsTmp.Fields("CardExpDate"), strCardExpDate)
                                    '#3531 �P�_�ϥΨ��@�خ榡 By Kin 2007/11/01
                Else
                    strErrMessage = rsTmp.Source
                End If
                begintrnsErrCode = 12.54
                file(1).WriteLine strErrMessage
                '#4090 �p�G���Ŀ�]�n���͸�ơA�p�G�S���Ŀ�h�������ܤU�@�� By Kin 2008/09/19
                If chkDuteDate.Value = 0 Then GoTo NextLoop
            End If
'' errcode**********
            begintrnsErrCode = 12.6
            '#3531 ���դ�OK�A���B��0���]�n��i�h By Kin 2008/05/22
            If rsTmp("ShouldAmt") <= 0 Then
                'strErrMessage = "�t�����B : �渹�@" & rsTmp.Fields("Billno") & " �Ȥ�m�W�@" & rsCustIDErr("custname") & "    ���B�@" & rsTmp("ShouldAmt")
                strErrMessage = "�t���άO���B����s : �渹�@" & rsTmp.Fields("Billno") & " �Ȥ�m�W�@" & rsCustIDErr("custname") & "    ���B�@" & rsTmp("ShouldAmt")
                file(1).WriteLine strErrMessage
                lngErrCount = lngErrCount + 1
                GoTo NextLoop
            End If

            begintrnsErrCode = 14
            If IsTCB = True Then
                Dim rsBudget As ADODB.Recordset
                Set rsBudget = GetBugetData(rsTmp("CitemCode"), rsTmp("ServiceType"), rsTmp("PTCode"), rsTmp("ShouldAmt"), IIf(IsNull(rsTmp("TCBbudget")), 0, rsTmp("TCBbudget")))
                If rsBudget.EOF Or rsBudget.BOF Then
                    lngErrCount = lngErrCount + 1
                    strErrMessage = "TCB�~�Ȭ����������O��  ������ƥ��]  " & _
                                              "Citemcode =" & rsTmp("CitemCode") & "�B" & _
                                              "ServiceType=" & rsTmp("ServiceType") & "�B" & _
                                              "PTCode=" & rsTmp("PTCode") & "�B" & _
                                              "ShouldAmt=" & rsTmp("ShouldAmt") & "�B" & _
                                              "TCBbudget=" & rsTmp("TCBbudget") & " --   " & _
                                              "�渹�@" & rsTmp.Fields("Billno") & " �Ȥ�m�W�@" & rsCustIDErr("custname") & " �H�Υd�� " & rsTmp.Fields("AccountNO")
                      file(1).WriteLine strErrMessage
                      GoTo NextLoop
                 End If
            End If
'' errcode**********
            begintrnsErrCode = 15
            lngCount = lngCount + 1
            '#3531 �p���`���B By Kin 2007/11/02
            lngShouldAmt = lngShouldAmt + Val(rsTmp("ShouldAmt"))
            With rsDefTmp
                .AddNew
                 ''  2003/10/08  �p�G�O����ӻȪ��ܧ飼�t�@�Ӱj��
                Select Case frmCreditCardType
                Case 0
                    If IsTCB = False Then
                        .Fields("F110") = GetString(txtSpcNO.Text, 10, giLeft, False)
                        .Fields("F1118") = GetString(txtMerchantName.Text, 8, giLeft, False)
                        .Fields("F1937") = GetString(rsTmp("AccountNO"), 19, giLeft, True)
                        .Fields("F3845") = GetString(rsTmp("ShouldAmt"), 8, GIRIGHT, True)
                         lngtotal = lngtotal + rsTmp("ShouldAmt")
                        .Fields("F4653") = Space(8)
                        .Fields("F5455") = "0" & cboSet.ListIndex    ''  ����X
                        .Fields("F5661") = Format(GiDAccept.Text, "EEMMDD")
                        If optmmyy.Value = True Then
                            '.Fields("F6265") = IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1)) & Right(rsTmp("CardExpDate"), 2)
                            .Fields("F6265") = Right(rsTmp("CardExpDate"), 2) & Mid(strCardExpDate, 3, 2)
                        Else
                            .Fields("F6265") = Mid(strCardExpDate, 3, 2) & Right(rsTmp("CardExpDate"), 2)
                            '.Fields("F6265") = Right(rsTmp("CardExpDate"), 2) & IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1))
                        End If
                        .Fields("F6668") = Space(3)
                        .Fields("F6984") = Space(16)
                        .Fields("F8590") = Space(6)
                        .Fields("F9196") = Space(6)
                       
                       
                        'If intPara24 = 0 Then
                        .Fields("F97112") = GetString(rsTmp("BillNO"), 16, giLeft, False)
                        
                       .Fields("F113118") = Space(6)
                       .Fields("F119") = chkCancel.Value
                       .Fields("F120") = cboMore.ListIndex
                    Else
                        Dim strTwDate As String
                        strTwDate = ""
                        strTwDate = Right((CInt(Right(rsTmp("CardExpDate"), 4)) - 1911), 2) & IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1))
                        If IsTCBCard = True Then
                            .Fields("F06") = Format(GiDAccept.Text, "EEMMDD")
                            .Fields("F10") = GetString(lngCount, 4, GIRIGHT, True)
                            .Fields("F26") = GetString(rsTmp("AccountNO"), 16, giLeft, False)
                            .Fields("F29") = GetString(rsTmp("CVC2") & "", 3, GIRIGHT, True)
                            .Fields("F33") = strTwDate
                            If intPara24 = 0 Then
                                .Fields("F44") = GetString(GetBillNo_New(rsTmp("BillNO")), 11, giLeft, False)
                            Else
                                .Fields("F44") = GetString(rsTmp("BillNo"), 11, giLeft, False)
                            End If
                            .Fields("F50") = Format(GiDAccept.Text, "EEMMDD")
                            .Fields("F56") = Space(6)
                            .Fields("F58") = GetString(rsBudget("Period"), 2, GIRIGHT, True) '' ����
                            .Fields("F65") = GetString(rsBudget("BudgetAmount"), 7, GIRIGHT, True) ''�������B
                            .Fields("F72") = GetString(rsBudget("Amount"), 7, GIRIGHT, True)  ''�`���B
                            .Fields("F79") = GetString(rsBudget("BudgetAmount"), 7, GIRIGHT, True)       ''�Y��
                            .Fields("F86") = GetString(rsBudget("LastAmount"), 7, GIRIGHT, True)       ''����
                            .Fields("F91") = "10000"
                            Dim strMerENG As String
                            Dim strMerCHI As String
                            strMerENG = txtSpcNO.Text & Space(50)
                            strMerCHI = txtSpcNO.Text & Space(80)
                            .Fields("F116") = StrConv(LeftB(StrConv(StrConv(strMerENG, vbWide), vbFromUnicode), 25), vbUnicode)
                            .Fields("F136") = StrConv(LeftB(StrConv(StrConv(strMerCHI, vbWide), vbFromUnicode), 20), vbUnicode)
                            
                            .Fields("F138") = "05"
                            .Fields("F180") = Space(62)
                            
                             lngtotal = lngtotal + rsTmp("ShouldAmt")   ' '  �O���`���B  ��b�ɧ�
                        Else
                            .Fields("F15") = GetString(rsTmp("BillNO"), 16, giLeft, False)
                            .Fields("F16") = GetString(rsTmp("BillNO"), 15, giLeft, False)
                            .Fields("F31") = GetString(rsTmp("AccountNO"), 19, giLeft, False)
                            .Fields("F50") = GetString(rsTmp("CVC2") & "", 3, giLeft, False)
                                         ''�אּ�褸�~      2003/05/05
                                         ''.Fields("F53") = strTwDate    ''�����
                            '.Fields("F53") = Right(rsTmp("CardExpDate"), 2) & IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1))
                             .Fields("F53") = Right(rsTmp("CardExpDate"), 2) & Mid(strCardExpDate, 3, 2)
                                         '' �`���B�אּ  10 �X
                                         '' .Fields("F57") = GetString(rsBudget("Amount"), 12, GIRIGHT, True) '' �`���B
                            .Fields("F57") = GetString(rsBudget("Amount"), 10, GIRIGHT, True) '' �`���B
                                          ''���ƹw�]�אּ   '00'
                            .Fields("F67") = "00"    '' �`�]��
                            .Fields("F69") = GetString(rsBudget("Period"), 2, GIRIGHT, True)  '' ����
                            .Fields("F71") = GetString(rsBudget("BudgetAmount"), 12, GIRIGHT, True)       ''�Y��
                            .Fields("F83") = GetString(rsBudget("LastAmount"), 12, GIRIGHT, True)       ''����
                             lngtotal = lngtotal + rsBudget("Amount")   ' '  �O���`���B  ��b�ɧ�
                        End If
                    End If
                Case 1    '' 2003/10/24 �H�U�o�@�q�O����ӻȪ��j��
                            ''''�S���N���G�̷~�Ȧ����P�N�� , �|���T�w
                            ''''�d���G�p�X�H�Υd�ɥd���u��11��: ����4�� '4000',�A��11��d��,�A��1��'0',�զ�16�� ; ��L�d��16��
                            ''''����X : ��l ��0 ��+�ť� , �g���v�B�z��|�ܬ� ��1��+�ť� , �j�����v(�G�����v)��|�ܬ� ��2��+�ť�
                            ''''�y����:  �ӵ����Ӧb�ɮפ��ĴX�� , ���藍�୫�_
                            ''''�^���X : ��1+�ť�+�ťա��Ρ� 5+�ť�+�ťա��ݦ��\���,��������^���T�����w��APPROVED ; �������~���^���X���ݱ��v���ѥ��
                            ''''�^���T�� : ��APPROVED���N���\ ; ��REFERRAL���N��CALL BANK,��L�h�N����,�̱`����TRANSACTION ERRO���N���ƿ��~ ; ��INVALID CAPTURE���N���L���� ; ��DECLINE���N���L����,�Բӭ�]�наѷӡ��妸�^���X���@��
                    begintrnsErrCode = 16.1
                    .Fields("F113") = GetString(txtMerchantName.Text, 13, giLeft, False)  ''  ** 13�S���N��(�����p�X:�ݥ����N��)
                      ''** 16�d��
                    If Len(Trim(rsTmp("AccountNO"))) = 11 Then
                        .Fields("F1429") = "4000" & rsTmp("AccountNO") & "0":          begintrnsErrCode = 16.2
                    Else
                        .Fields("F1429") = GetString(rsTmp("AccountNO"), 16, giLeft, True):           begintrnsErrCode = 16.3
                    End If
                    .Fields("F3039") = GetString(rsTmp("ShouldAmt"), 10, GIRIGHT, True)  ''  ** 10���B
'                    .Fields("F4043") = IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2),  "0" & Left(rsTmp("CardExpDate"), 1)) & Right(rsTmp("CardExpDate"), 2 )   ''   ** 4���Ĥ�~ (MMYY) (�褸�~)
                    .Fields("F4043") = Right(rsTmp("CardExpDate"), 2) & Mid(strCardExpDate, 3, 2)
                    .Fields("F4445") = "0" & Space(1)  '' ** 2����X   (0+�ť�)
                    ''                       .Fields("F4659") = GetString(lngCount, 14, GIRIGHT, True) ''** 14�y����(���藍�i����)  (������0)
                    If intPara24 = 0 Then
                        .Fields("F4659") = GetString(GetBillNo_New(rsTmp("BillNO")), 14, GIRIGHT, True)  ''** 2003/10/13 �o�����令��渹
                    Else
                        .Fields("F4659") = GetString(rsTmp("BillNO"), 14, GIRIGHT, True)  ''** 2003/10/13 �o�����令��渹
                    End If
                    '' �H�U���Ȧ�^�ХΪ�
                    .Fields("F6065") = Space(6)
                    .Fields("F6671") = Space(6)
                    .Fields("F7274") = Space(3)
                    .Fields("F7590") = Space(16)
                    .Fields("F9196") = Space(6)
                    .Fields("F97102") = Space(6)
                    lngtotal = lngtotal + rsTmp("ShouldAmt")
                Case 2    '' 2003/11/20 �H�U�o�@�q�O���׻Ȧ檺�j��
:                           begintrnsErrCode = 16.4
                    .Fields("F116") = GetString(rsTmp("AccountNO"), 16, giLeft, False)    ''�d��  �����A�̨t�γW�w�ɨ�16�Xspace
                    .Fields("F1717") = Space(1)
                    begintrnsErrCode = 16.5
                    .Fields("F1821") = IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1)) & Right(rsTmp("CardExpDate"), 2)      ''  ** ���Ĵ���  MBF,MMYY
                    .Fields("F2222") = Space(1):   begintrnsErrCode = 16.6
                    .Fields("F2334") = GetString(rsTmp("ShouldAmt") & "00", 12, GIRIGHT, True)   ''  ** ���B  ������'0'���X������
                    .Fields("F3535") = Space(1): begintrnsErrCode = 16.7
                    
                    .Fields("F3641") = "999999"  ''   �ֲa���X   ��J'999999'or approvaled code
                    .Fields("F4242") = Space(1): begintrnsErrCode = 16.8
                    .Fields("F4344") = "99"   ''  �^�_�X   ��J'99'or '0'(if this is a refund)
                    .Fields("F4545") = Space(1): begintrnsErrCode = 16.9
                    .Fields("F4685") = GetString(rsTmp("BillNO"), 40, giLeft, False)     '' �ө��ۦ�B�� (so043.para24=1�h���C��渹
                    '' �p��������ƥH�Φ��O���B
                    If rsTmp("ShouldAmt") > 0 Then
                        lngTotalHSBC = lngTotalHSBC + rsTmp("ShouldAmt"): begintrnsErrCode = 17.3
                        lngToatalCountHSBC = lngToatalCountHSBC + 1: begintrnsErrCode = 17.2
                    Else
                        lngTotalHSBCN = lngTotalHSBCN + Abs(rsTmp("ShouldAmt")): begintrnsErrCode = 17
                        lngTotalCountHSBCN = lngTotalCountHSBCN + 1: begintrnsErrCode = 17.1
                    End If
                '3531 �ϥηs�榡 By Kin 2007/11/01
                Case 3
                    .Fields("SPCNO") = txtAuthorizationBatch.Text
                    .Fields("ACCOUNTID") = rsTmp("ACCountNO")
                    '���~�����X��10/2010 ��1010
                    .Fields("STOPYM") = Trim(Right(strCardExpDate, 2)) & Mid(strCardExpDate, 3, 2)      ' Trim(Str(Val(Left(strCardExpDate, 4)) - 1911))
                    .Fields("SHOULDAMT") = rsTmp("ShouldAmt")
                    .Fields("BILLNO") = rsTmp("BillNo")
                    .Fields("TRADETYPE") = Left(cboMore.Text, 1)
                    '#4033 �쥻��101~140�����r��,�אּ101~138�᭱��2�Ӫť� By Kin 2008/08/14
                    .Fields("CHINESE") = GetString(StrConv(txtBillMemo.Text, vbWide) & StrConv(Space(38), vbWide), 38, giLeft) & Space(2)
                    
                '#5494 �W�[�ŷs��ޮ榡 By Kin 2010/01/27
                Case 5
                    .Fields("RcdType") = "D"
                    .Fields("OrderNo") = GetString(rsTmp("BillNo") & "", 30, giLeft)
                    .Fields("OrdItem") = Space(9)
                    '��h�W�������Ƥ@�w�|�O�ť�
                    If rsTmp("TCBbudget") & "" = "" Then
                        .Fields("Period") = Space(2)
                    Else
                        .Fields("Period") = GetString(rsTmp("TCBbudget") & "", 2, GIRIGHT, True)
                    End If
                    .Fields("Status") = Space(1)
                    .Fields("Amt") = GetString(FmtAmt(rsTmp("ShouldAmt") & ""), 13, GIRIGHT, True)
                    .Fields("CardNo") = GetString(rsTmp("AccountNo") & "", 16, GIRIGHT, True)
                    If .Fields("CVC2") & "" <> "" Then
                        .Fields("CVC2") = GetString(rsTmp("CVC2") & "", 3, GIRIGHT, True)
                    Else
                        .Fields("CVC2") = Space(3)
                    End If
                    .Fields("ValidDate") = GetString(Replace(strCardExpDate, "/", ""), 6, GIRIGHT, True)
                    .Fields("AuthNo") = Space(6)
                    .Fields("Prc") = Space(5)
                    .Fields("Src") = Space(5)
                    .Fields("RetBank") = Space(2)
                    .Fields("BatNum") = Space(9)
                    .Fields("AutoFlag") = "1"
                    .Fields("hold") = Space("91")
                '#5609 �W�[�p�X�H�Υd�妸�榡 By Kin 2010/06/07
                Case 6
'                    .Fields.Append "SPCNO", adBSTR, 10, adFldIsNullable         '�ө��N�X
'                    .Fields.Append "MERCHANTNAME", adBSTR, 8, adFldIsNullable   '�ݥ����N��
'                    .Fields.Append "ACCOUNTNO", adBSTR, 16, adFldIsNullable     '�d��
'                    .Fields.Append "ACCOUNTNOEX", adBSTR, 3, adFldIsNullable    '�d�������X
'                    .Fields.Append "STOPYM", adBSTR, 4, adFldIsNullable         '���Ĵ���
'                    .Fields.Append "AMT", adBSTR, 10, adFldIsNullable           '������B
'                    .Fields.Append "AUTHNO", adBSTR, 8, adFldIsNullable         '���v�X
'                    .Fields.Append "SETNO", adBSTR, 2, adFldIsNullable          '����X
'                    .Fields.Append "ACCEPTDATE", adBSTR, 6, adFldIsNullable     '�e����
'                    .Fields.Append "BILLNO", adBSTR, 16, adFldIsNullable        'USER DEFINE
'                    .Fields.Append "ADDITION", adBSTR, 30, adFldIsNullable      'Additional field
'                    .Fields.Append "ACCINFO", adBSTR, 40, adFldIsNullable       '�d�H��T
                    .Fields("SPCNO") = GetString(txtSpcNO, 10, giLeft)
                    .Fields("MERCHANTNAME") = GetString(txtMerchantName.Text, 8, giLeft, False)
                    .Fields("ACCOUNTNO") = GetString(rsTmp("AccountNO"), 16, giLeft, True)
                    .Fields("ACCOUNTNOEX") = "000"
'                    .Fields("STOPYM") = Right(rsTmp("CardExpDate"), 2) & IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1))
                    .Fields("STOPYM") = Right(Left(rsTmp("CardExpDate") & "", 4), 2) & Right(rsTmp("CardExpDate") & "", 2)
                    .Fields("AMT") = GetString(rsTmp("ShouldAmt"), 10, GIRIGHT, True)
                    .Fields("AUTHNO") = GetString(rsTmp("CVC2") & "", 8, giLeft)
                    .Fields("SETNO") = Left(cboSet.Text, 2)
                    .Fields("ACCEPTDATE") = Format(GiDAccept.Text, "EEMMDD")
                    .Fields("BILLNO") = GetString(rsTmp("BillNo") & "", 16, giLeft)
                    .Fields("ADDITION") = Space(30)
                    If txtBillMemo.Text = "" Then
                        .Fields("ACCINFO") = String(40, StrConv(" ", vbWide))
                    Else
                        .Fields("ACCINFO") = Left(StrConv(txtBillMemo.Text, vbWide) & _
                                        String(40, StrConv(" ", vbWide)), 40)
                    End If
                        
                End Select
                    .Update
            End With
                        
            '**********************************************************************************************************************************
            '#3417 �p�GCD013.REFNO���]�t��4,�h��sUCCode�PUCName By Kin 2007/12/05
            '#4388 �W�[��s���ʤH���P���ʮɶ� By Kin 2009/04/30
            '#5218 �d�x�w���n��Nvl By Kin 2009/08/05
            If blnUpdUCCode Then
              
                '#3892 ��s�t���ܺC,�վ�y�k,�쥻��Exists,�{�b���= By Kin 2008/06/03
                '#4037 �|�N�Ҧ��۲Ū��C�^�渹��Upd��,�令RowId In(...) By Kin 2008/09/11
                '#5267 WHERE����SO002A���SO106 BY Kin 2010/04/20
                '#5564 �ѦҸ�7,8�PCD013.PAYOK=1���O�N��w�� By Kin 2010/05/17
                strUpdUCCode = "UPDATE " & TableOwnerName & "SO033  Set UCCode=" & strUCCode & ",UCName='" & strUCName & "'" & _
                                ",UPDEN='" & objStorePara.uUPDEN & "',UPDTIME='" & objStorePara.uUPDTime & "' " & _
                            " Where RowId  In (Select A.RowId From " & _
                                            TableOwnerName & "SO001 B ," & _
                                            TableOwnerName & "SO002 D," & _
                                            TableOwnerName & "SO106 C," & _
                                            TableOwnerName & "SO033 A," & _
                                            TableOwnerName & "CD013 " & _
                                            " Where " & strUCCodeWhere & _
                                            " And A." & strMediaBillNOField & "='" & rsTmp("BillNo") & "'" & _
                                            " And A.AccountNO='" & rsTmp("AccountNO") & "'" & _
                                            " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN (3,7) AND  " & _
                                            " NVL(CD013.PAYOK,0) = 0 " & _
                                            IIf(rsTmp("CardExpDate") & "" <> "", " And " & _
                                            "DECODE(LENGTH(C.STOPYM),5,SUBSTR(C.STOPYM,2,4) || '0' || SUBSTR(C.STOPYM,1,1),SUBSTR(C.STOPYM,3,4)||SUBSTR(C.STOPYM,1,2))  = " & rsTmp("CardExpDate"), "") & _
                                            IIf(rsTmp("CVC2") & "" <> "", " And C.CVC2=" & rsTmp("CVC2"), "") & ")"
                cn.Execute strUpdUCCode
            End If
            '**********************************************************************************************************************************
            
NextLoop:
            rsTmp.MoveNext
        Wend
        '' �H�U���o header �r��
        Select Case frmCreditCardType
        Case 0
'            strHeader = "Header"
'            strHeader = strHeader & GetString(lngCount, 8, GIRIGHT, True)
'            strHeader = strHeader & Space(106)
            strHeader = ""
        Case 1
            ''' 2003/10/08 �H�U�o�@�q�O����ӻȪ��j��
            strHeader = ""
            strHeader = GetString(txtSpcNO.Text, 6, giLeft, False) & Format(Date, "eemmdd") & _
                        GetString(lngCount, 10, GIRIGHT, True) & GetString(lngtotal, 12, GIRIGHT, True) & _
                        Space(10) & Space(12) & Space(10) & Space(12)
        Case 2
            begintrnsErrCode = 17.5
            strHeader = "": begintrnsErrCode = 17.6
            strHeader = GetString(Left(txtMerchantName.Text, 8), 8, giLeft, False) & Space(1) & _
                        GetString(Left(txtSpcNO.Text, 15), 15, giLeft, False) & Space(1) & _
                        GetString(lngTotalCountHSBCN + lngToatalCountHSBC, 5, GIRIGHT, True) & Space(1) & _
                        GetString((lngTotalHSBC + lngTotalHSBCN) & "00", 12, GIRIGHT, True) & Space(1) & _
                        GetString(lngTotalCountHSBCN, 5, GIRIGHT, True) & Space(1) & _
                        GetString((lngTotalHSBCN) & "00", 12, GIRIGHT, True) & Space(1) & _
                        GetString((lngTotalHSBC - lngTotalHSBCN) & "00", 12, GIRIGHT, True)
        '#3531 ��J���Y By Kin 2007/11/01
        Case 3
            '#3531 ���դ�OK,�n��J183��0�A����O�ť� By Kin 2007/11/22
            strHeader = ""
            strHeader = "H" & GetString(txtSpcNO, 15) & _
                        "E" & Space(183)
        '#5609 �W�[�p�X�H�Υd�妸�榡 By Kin 2010/06/03
        Case 6
            strHeader = ""
            strHeader = "HEADER" & Format(GiDAccept.Text, "EEMMDD") & _
                        GetString(lngCount, 10, GIRIGHT, True) & _
                        GetString(lngtotal, 12, GIRIGHT, True) & _
                        Space(46)
            
        End Select
        '' 2003/01/23  �p�G�O���ذӻȪ� �b�o�@����ɧ����O��  ........
        If IsTCB = True Then
             Dim strEndLine As String
             strEndLine = "T"
             strEndLine = strEndLine & GetString(lngCount, 6, GIRIGHT, True)
             
             ''  �`���B�p�G�O�L��d�A�אּ �Q �X �A�]���H�U�o�@����O�̪��פ��P�אּ�b    If IsTCBCard = True Then  ���϶����Y�@
            ''strEndLine = strEndLine & GetString(lngtotal, 12, GIRIGHT, True)
            If IsTCBCard = True Then
                strEndLine = strEndLine & GetString(lngtotal, 12, GIRIGHT, True)
                strEndLine = strEndLine & Format(Date, "EEMMDD")
            Else
                strEndLine = strEndLine & GetString(lngtotal, 10, GIRIGHT, True)
                strEndLine = strEndLine & Format(Date, "YYYYMMDD")
            End If
            strEndLine = strEndLine & Format(Now, "HHMMSS")
            If IsTCBCard = False Then
                Dim rsCD018Tmp As New ADODB.Recordset
                Dim strNN As String
                strNN = ""
                '                                                     rsCD018Tmp.CursorLocation = adUseClient
                '                                                     rsCD018Tmp.Open "SELECT  *  FROM " & TableOwnerName & "CD018Tmp   ORDER BY BATCHNODATE  DESC , BATCHNO DESC  ", cn, adOpenKeyset, adLockReadOnly
                If Not GetRS(rsCD018Tmp, "SELECT  *  FROM " & TableOwnerName & "CD018Tmp   ORDER BY BATCHNODATE  DESC , BATCHNO DESC  ", cn) Then Exit Function
                If rsCD018Tmp.EOF Or rsCD018Tmp.BOF Then
                    strNN = "01"
                Else
                    If Format(Date, "yyyymmdd") = rsCD018Tmp("BatchNODate") Then
                        If rsCD018Tmp("BatchNO") < 9 Then
                            strNN = "0" & CStr(rsCD018Tmp("BatchNO") + 1)
                        Else
                            strNN = CStr(rsCD018Tmp("BatchNO") + 1)
                        End If
                    Else
                        strNN = "01"
                    End If
                End If
                rsCD018Tmp.Close
                Set rsCD018Tmp = Nothing
                strEndLine = strEndLine & Left(txtSpcNO.Text, 4) & "AU" & Format(Date, "YYMMDD") & strNN
                If lngCount > 0 Then
                    cn.Execute "INSERT INTO CD018Tmp (BatchNO, BatchNODate) values (" & CInt(strNN) & ",'" & Format(Date, "yyyymmdd") & "')"
                End If
            End If
        End If
        '#3531 ����H�U�s�榡�ݭn��J�s�榡 By Kin 2007/11/02
        If frmCreditCardType = lngCtrust2 Then
            Dim intSpcCount As Integer
            Dim intKeep As Integer
            strEndLine = ""
            strEndLine = "R001"
            intSpcCount = 18
            intKeep = 142 '�O�d���
            Select Case Left(cboMore.Text, 1)
                Case "S", "O"
                    strEndLine = strEndLine & GetString(CStr(lngCount), 5, GIRIGHT, True) & _
                                GetString(CStr(lngShouldAmt), 11, GIRIGHT, True) & "00" & String(36, "0") & Space(intKeep)
                Case "A"
                    strEndLine = strEndLine & String(intSpcCount, "0") & _
                                GetString(CStr(lngCount), 5, GIRIGHT, True) & _
                                GetString(CStr(lngShouldAmt), 11, GIRIGHT, True) & "00" & String(intSpcCount, "0") & Space(intKeep)
                Case "R"
                    strEndLine = strEndLine & String(intSpcCount * 2, "0") & _
                                GetString(CStr(lngCount), 5, GIRIGHT, True) & _
                                GetString(CStr(lngShouldAmt), 11, GIRIGHT, True) & "00" & Space(intKeep)
            End Select
        End If
        '#5494 �ŷs��ު����Y By Kin 2010/02/02
        If frmCreditCardType = lngNewWeb Then
            strHeader = "H" & GetString(GiDAccept.GetValue, 8, GIRIGHT, True) & _
                GetString(txtAuthorizationBatch.Text, 2, GIRIGHT, True) & _
                GetString(rsDefTmp.RecordCount, 8, GIRIGHT, True) & _
                GetString(FmtAmt(CStr(lngShouldAmt)), 15, GIRIGHT, True) & Space(166)
                


        
        End If
    End If
        begintrnsErrCode = 17.7
        If IsTCB = True Then
            Call subWriteLine(strEndLine)
        Else
            Call subWriteLine(strEndLine)
        End If
            begintrnsErrCode = 17.8
        msgResult lngCount, lngErrCount, lngTime      '��ܰ��浲�G
       
        BeginTran = True
        On Error Resume Next
        If rsSO106.State = adStateOpen Then rsSO106.Close
        If rsCD013.State = adStateOpen Then rsCD013.Close
        If rsUpdUCCode.State = adStateOpen Then rsUpdUCCode.Close
        rsTmp.Close
        Set rsTmp = Nothing
        Set rsSO106 = Nothing
        Set rsUpdUCCode = Nothing
        Set rsCD013 = Nothing
        
        
        
        CloseFS
     
    If lngDoubleMediaNO > 0 Then
    '     Call MsgBox("�X�檺��Ʀ��ǬO���Ъ��A�Ьd�\�O���ɽT�{�һݪ���Ƥ��e�A�קK���ХX��  !!", vbExclamation + vbOKOnly, "���ХX��ĵ��")
    '     Call Shell("NotePad " & txtErrPath.Text, vbNormalFocus)
    End If
    
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "BeginTran �GerrCode = " & begintrnsErrCode & "---" & strBankId & ";ErrNumber:" & Err.Number & ";ErrDescription:" & Err.Description)
    CloseFS
'    MsgBox begintrnsErrCode
'    MsgBox "�渹�@" & rsTmp.Fields("Billno") & " �Ȥ�m�W�@" & rsCustIDErr("custname") & " �H�Υd�� " & rsTmp.Fields("AccountNO")
    
End Function
'#5494 �ŷs�榡�����B��2�X�O�p���I
Private Function FmtAmt(ByVal aShouldAmt As String) As String
  On Error GoTo ChkErr
    Dim sTmp As String
    If InStr(1, aShouldAmt, ".") Then
        sTmp = Left(aShouldAmt, InStr(1, aShouldAmt, "."))
        sTmp = sTmp & GetString(Mid(aShouldAmt, InStr(1, aShouldAmt, ".") + 1, 2), 2, giLeft, True)
        sTmp = Replace(sTmp, ".", "")
    Else
        sTmp = aShouldAmt & "00"
    End If
    FmtAmt = sTmp
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "FmtAmt")
End Function

Private Function subWriteLine(Optional pEndLine As String = "") As Boolean

  On Error GoTo ChkErr
  
    Dim varData As Variant
    Dim strData As String
    Dim lngLoop As Long
    Dim lngLoopi As Long
    Dim strContent As String

    subWriteLine = False
    
    With rsDefTmp
        If .BOF Or .EOF Then Exit Function
            .MoveFirst
         
            If IsTCB = False Then
                If strHeader <> Empty Then
                    file(0).WriteLine (strHeader)    ''���r�ɶ�J����
                End If
            End If
         
            For lngLoop = 0 To .RecordCount - 1
                 strContent = ""
                 Dim i As Long
                 
                 ''2003/10/24  �ھګH�Υd�O�A�h�]�w������榡
                Select Case frmCreditCardType
                    Case 0
                        If IsTCB = False Then
                            i = 15
                        Else
                            If IsTCBCard = True Then i = 17 Else i = 8
                        End If
                    Case 1
                        i = 11
                    Case 2
                        i = 10
                    '#5494 �ŷs��� By Kin 2010/02/03
                    '#5609 �p�X�H�Υd�妸�榡 By Kin 2010/06/07
                    Case 5, 6
                        i = rsDefTmp.Fields.Count - 1
                    
                End Select
                '#3531 �s�W����H�U�s�榡(����) By Kin 2007/11/02
                If frmCreditCardType = lngCtrust2 Then
                    strContent = "T" & GetString(.Fields("SPCNO") & "", 5, GIRIGHT, True) & _
                                    GetString(.Fields("ACCOUNTID") & "", 19) & .Fields("STOPYM") & _
                                    GetString(.Fields("ShouldAMT") & "", 11, GIRIGHT, True) & "00" & _
                                    Space(2) & GetString(.Fields("BillNo") & "", 15) & Space(6) & _
                                    .Fields("TRADETYPE") & "" & Space(34) & .Fields("CHINESE") & Space(60)
                                        
                Else
                    For lngLoopi = 0 To i
                        strContent = strContent & rsDefTmp(lngLoopi)
                    Next
                End If
                file(0).WriteLine (strContent)
                rsDefTmp.MoveNext
                DoEvents
            Next
            '#3531 �p�G�O����H�U�s�榡�]�n�g�J��� By Kin 2007/11/02
            If IsTCB = True Or frmCreditCardType = lngCtrust2 Then
                file(0).WriteLine (pEndLine)
            End If
            
    End With
    subWriteLine = True
    On Error Resume Next
    rsDefTmp.Close
    Set rsDefTmp = Nothing
    
    Exit Function
ChkErr:
    subWriteLine = False
    Call ErrSub(Me.Name, "subWriteLine")
End Function

Private Sub ScrToRcd()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    Set LogFile = fso.CreateTextFile(strPath & "\" & strPrgName & "Out.log", True)
    '#5494 �h�W�[�C�릩�ڤ� By Kin 2010/02/03
    With LogFile
        .WriteLine (txtSpcNO.Text) '�H�Υd���b�S���W��
        .WriteLine (txtMerchantName.Text) '�H�Υd���b�S���W��
        .WriteLine (txtDataPath)         '����ɸ��|
        .WriteLine (txtErrPath.Text)     ' ���D�Ѧ��ɸ��|
        .WriteLine (cboMonthDay.Text)
    End With
        
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub
Private Function IsDataOk() As Boolean
    Dim strerrmsg As String
    
    If IsTCB = False Then
        Select Case frmCreditCardType
            Case 0
                If ChkFields(True, True, True, True) = False Then Exit Function
            Case 1
                If ChkFields(True, True, False, False) = False Then Exit Function
            Case 2
                If ChkFields(True, True, False, False, True) = False Then Exit Function
            Case 3
                If ChkFields(True, False, False, True, True) = False Then Exit Function
            '#5494 �W�[�ŷs��ު��P�_ By Kin 2010/02/03
            Case 5
                If ChkFields(True, False, False, False, False, True) = False Then Exit Function
            '#5609 �W�[�p�X�H�Υd�妸�ˬd By Kin 2010/06/03
            Case 6
                If ChkFields(True, True, True, True) = False Then Exit Function
        End Select
    End If

    If Len(GiDAccept.Text) = 0 Then strerrmsg = "�e����": GiDAccept.SetFocus: GoTo Warning
    If Len(txtDataPath.Text) = 0 Then strerrmsg = "����ɸ��|": txtDataPath.SetFocus: GoTo Warning
    If Len(txtErrPath.Text) = 0 Then strerrmsg = "���D�Ѧ��ɸ��|": txtErrPath.SetFocus: GoTo Warning
    IsDataOk = True
    
    Exit Function
Warning:
    MsgBox "��� " & strerrmsg & " ���ݦ���", vbOKOnly + vbInformation, "�T��"
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function
Private Function ChkFields(ByVal blnSpcNO As Boolean, ByVal blnMerchantName As Boolean, _
                           ByVal blnSet As Boolean, ByVal blnMore As Boolean, _
                           Optional blnAuthorizationBatch As Boolean = False, _
                           Optional ByVal blnAcceptDate As Boolean = False) As Boolean
                                   
    Dim strerrmsg As String
    ChkFields = False
    If blnSpcNO = True Then
        If Len(txtSpcNO.Text & "") = 0 Then strerrmsg = "�ө��N��": txtSpcNO.SetFocus: GoTo Warning
    End If
    If blnMerchantName = True Then
        If Len(txtMerchantName.Text & "") = 0 Then strerrmsg = "�ݥ����N��": txtMerchantName.SetFocus: GoTo Warning
    End If
    If blnSet = True Then
        If Len(cboSet.Text) = 0 Then strerrmsg = "����X": cboSet.SetFocus: GoTo Warning
    End If
    If blnMore = True Then
        If Len(cboMore.Text) = 0 Then strerrmsg = "����X��": cboMore.SetFocus: GoTo Warning
    End If
    If blnAuthorizationBatch = True Then
        If Len(txtAuthorizationBatch.Text) = 0 Then strerrmsg = "���v�妸": txtAuthorizationBatch.SetFocus: GoTo Warning
    End If
    If blnAcceptDate Then
        If Len(GiDAccept.GetValue) = 0 Then strerrmsg = "�e����": GiDAccept.SetFocus: GoTo Warning
    End If
    
    ChkFields = True

Exit Function
Warning:
    MsgBox "��� " & strerrmsg & " ���ݦ���", vbOKOnly + vbInformation, "�T��"
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call CloseFS
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
End Sub

Private Sub GiDAccept_GotFocus()
    If Len(GiDAccept.Text) = 0 Then GiDAccept.Text = Format(Date, "EE/MM/DD")
End Sub

Private Sub txtAuthorizationBatch_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    If KeyAscii = 8 Then Exit Sub
    If frmCreditCardType = 3 Or frmCreditCardType = 5 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    End If
    KeyAscii = 0
End Sub

Private Sub txtMerchantName_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    
End Sub

Public Function GetInvoiceNo3(StrTableName As String, Optional Conn As ADODB.Connection) As String
  On Error GoTo ChkErr
    Dim strSeq As String
    Dim strInv
    strSeq = "S_" & StrTableName
    ''If Conn Is Nothing Then Set Conn = gcnGi
    strInv = GetRsValue("SELECT Ltrim(To_Char(" & strSeq & ".NextVal, '09999999')) FROM Dual", Conn)
    GetInvoiceNo3 = strInv
    Exit Function
ChkErr:
  ErrSub Me.Name, "GetInvoiceNo3"
End Function
Public Function BatchUpdateMediaBillNo(rsSource As ADODB.Recordset, strChoose33 As String, _
cn As ADODB.Connection) As Boolean
'On Error GoTo ChkErr
'Dim lngAffected As Long
'Dim strMediaBillNo As String
'        With rsSource
'            If .RecordCount > 0 Then .MoveFirst
'            Do While Not .EOF
'                    strMediaBillNo = ""
'                    If Not GetMediaBillNo(rsSource("CustId"), strMediaBillNo, cn) Then Exit Function
'                    If Not ExecuteCommand("Update SO033 A Set MediaBillNo = '" & strMediaBillNo & "' Where BillNo = '" & rsSource("BillNo") & "' And CustId = " & rsSource("CustId") & IIf(Len(strChoose33) = 0, "", " And ") & strChoose33, cn, lngAffected) Then Exit Function
'                    On Error Resume Next
'                    .Fields("MediaBillNo") = strMediaBillNo
'                    .Update
'                    On Error GoTo ChkErr
'                    .MoveNext
'            Loop
'        End With
'      BatchUpdateMediaBillNo = True
'Exit Function
'ChkErr:
'            ErrSub FormName, "BatchUpdateMediaBillNo"
End Function

Public Function GetMediaBillNo(ByRef lngCustId As Long, ByRef strMediaBillNo As String, cn As ADODB.Connection) As Boolean
'''On Error GoTo ChkErr

'lngCustId = Val(GetInvoiceNo3("SO033_MediaBillNo", cn) & "")
'strMediaBillNo = Format(Date, "yymm") & Format(lngCustId, "0000000")
'GetMediaBillNo = True
'Exit Function
'ChkErr:
'ErrSub FormName, "GetMediaBillNo"
End Function

Private Function GetBugetData(ByVal intCitemCODE As Integer, _
                                                  ByVal strServiceType As String, _
                                                  ByVal intPTCode As Integer, _
                                                  ByVal intAmount As Integer, _
                                                  ByVal intPeriod As Integer) As ADODB.Recordset

    Dim strsql As String
    Dim rs As New ADODB.Recordset
  On Error GoTo ChkErr
        strsql = "SELECT * From " & TableOwnerName & "CD019B Where " & _
                    "Citemcode=" & intCitemCODE & " AND  ServiceType='" & strServiceType & "' AND " & _
                    "PTCode =" & intPTCode & " AND  Amount = " & intAmount & " AND Period = " & intPeriod
        If rs.State = 1 Then rs.Close
    '    rs.CursorLocation = adUseClient
    '    rs.Open strsql, cn, adOpenKeyset, adLockReadOnly
        If Not GetRS(rs, strsql, cn) Then Exit Function
        Set GetBugetData = rs
Exit Function
ChkErr:
    ErrSub Me.Name, "GetBugetData"
End Function

Private Function CheckOutDate() As Boolean

    Dim aPath() As String
    Dim varAnayPath
    Dim strAnayPath As String

  On Error GoTo ChkErr

    If frmCreditCardType = 2 Then
        aPath = Split(txtDataPath.Text, "\")
        For Each varAnayPath In aPath
            strAnayPath = varAnayPath
        Next
        aPath = Split(strAnayPath, ".")
        strAnayPath = aPath(0)
        CheckOutDate = True
        If Right(strAnayPath, 6) <> Format(Date, "eemmdd") Then
            MsgBox "��X�ɦW������D����A�жi��վ�    !!   "
            CheckOutDate = False
        Else
            If intBatchNumber = CInt(Mid(strAnayPath, 2, Len(strAnayPath) - 7)) Then
                MsgBox "��X�ɧ妸���СA�нվ��ɦW�妸���X    !!   "
                CheckOutDate = False
            End If
        End If
    Else
        CheckOutDate = True
    End If

    Exit Function
ChkErr:
    ErrSub Me.Name, "CheckOutDate"
End Function

