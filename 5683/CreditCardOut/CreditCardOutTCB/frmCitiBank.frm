VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCitiBank 
   BorderStyle     =   1  '��u�T�w
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCitiBank.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   9240
   StartUpPosition =   1  '���ݵ�������
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   5520
      Top             =   180
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
      Height          =   3975
      Left            =   165
      TabIndex        =   16
      Top             =   15
      Width           =   8895
      Begin VB.TextBox txtBatch 
         Height          =   345
         Left            =   3570
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "01"
         Top             =   1410
         Width           =   900
      End
      Begin VB.ComboBox cboBillType 
         ForeColor       =   &H000000FF&
         Height          =   315
         ItemData        =   "frmCitiBank.frx":0442
         Left            =   2040
         List            =   "frmCitiBank.frx":044F
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   0
         Top             =   240
         Width           =   1665
      End
      Begin VB.CheckBox chkDuteDate 
         Caption         =   "�H�Υd�L����Ƥ@�ֲ���"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2040
         TabIndex        =   9
         Top             =   2265
         Width           =   2535
      End
      Begin VB.TextBox txtSpcNO 
         Height          =   300
         Left            =   2025
         MaxLength       =   15
         TabIndex        =   1
         Top             =   705
         Width           =   1830
      End
      Begin VB.TextBox txtMerchantName 
         Height          =   300
         Left            =   6060
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "TBGB"
         Top             =   705
         Width           =   2460
      End
      Begin VB.TextBox txtStatement 
         Height          =   345
         Left            =   2025
         TabIndex        =   8
         Top             =   1800
         Width           =   6465
      End
      Begin VB.TextBox txtDiscountRate 
         Height          =   345
         Left            =   2025
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "0"
         Top             =   1410
         Width           =   900
      End
      Begin VB.TextBox txtPayDate 
         Height          =   345
         Left            =   6075
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "00"
         Top             =   1410
         Width           =   330
      End
      Begin VB.ComboBox cboSet 
         Height          =   315
         ItemData        =   "frmCitiBank.frx":047D
         Left            =   2025
         List            =   "frmCitiBank.frx":048D
         TabIndex        =   3
         Top             =   1050
         Width           =   2460
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
         Height          =   1275
         HelpContextID   =   2
         Left            =   285
         TabIndex        =   17
         Top             =   2475
         Width           =   8160
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
            TabIndex        =   10
            ToolTipText     =   "�п�J�r�ɤ����|���ɦW�I"
            Top             =   300
            Width           =   4095
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
            Left            =   7320
            TabIndex        =   13
            Top             =   780
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
            Left            =   7320
            TabIndex        =   11
            Top             =   315
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
            Left            =   3015
            TabIndex        =   12
            Top             =   765
            Width           =   4095
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "����ɦ�m�]���| )"
            Height          =   195
            Left            =   1290
            TabIndex        =   19
            Top             =   375
            Width           =   1665
         End
         Begin VB.Label lblErrpath 
            AutoSize        =   -1  'True
            Caption         =   "���D�Ѧ��ɦ�m�]���| + �W�١^"
            Height          =   195
            Left            =   330
            TabIndex        =   18
            Top             =   810
            Width           =   2640
         End
      End
      Begin Gi_Date.GiDate gdaSystemDate 
         Height          =   345
         Left            =   6060
         TabIndex        =   6
         Top             =   1035
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
      Begin VB.Label lblBatch 
         Alignment       =   1  '�a�k���
         AutoSize        =   -1  'True
         Caption         =   "�帹"
         Height          =   195
         Left            =   3015
         TabIndex        =   29
         Top             =   1500
         Width           =   390
      End
      Begin VB.Label Label6 
         Caption         =   "��X�榡"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1170
         TabIndex        =   28
         Top             =   330
         Width           =   825
      End
      Begin VB.Label lblSpcNO 
         Alignment       =   1  '�a�k���
         AutoSize        =   -1  'True
         Caption         =   "�ө��N��"
         Height          =   195
         Left            =   1155
         TabIndex        =   27
         Top             =   780
         Width           =   780
      End
      Begin VB.Label Label5 
         Alignment       =   1  '�a�k���
         AutoSize        =   -1  'True
         Caption         =   "�t�Τ��"
         Height          =   195
         Left            =   5190
         TabIndex        =   26
         Top             =   1155
         Width           =   780
      End
      Begin VB.Label Label9 
         Caption         =   "(DD)"
         Height          =   240
         Left            =   6465
         TabIndex        =   25
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label lblMerchantName 
         AutoSize        =   -1  'True
         Caption         =   "�ݥ����N��"
         Height          =   195
         Left            =   5025
         TabIndex        =   24
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�Ȥ�b�椺�e"
         Height          =   195
         Left            =   750
         TabIndex        =   23
         Top             =   1890
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   1  '�a�k���
         AutoSize        =   -1  'True
         Caption         =   "�馩�v"
         Height          =   195
         Left            =   1350
         TabIndex        =   22
         Top             =   1515
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "�дڤ��"
         Height          =   210
         Left            =   5190
         TabIndex        =   21
         Top             =   1515
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "���b����"
         Height          =   195
         Left            =   1140
         TabIndex        =   20
         Top             =   1155
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�}�l"
      Height          =   375
      Left            =   315
      TabIndex        =   14
      Top             =   4140
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   7665
      TabIndex        =   15
      Top             =   4125
      Width           =   1275
   End
End
Attribute VB_Name = "frmCitiBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Private cn As New ADODB.Connection
Private rsTmp As New ADODB.Recordset     '' ���X�����Ʈw �����
Private rsDefTmp As New ADODB.Recordset    ''�s��r�ɪ��O���� recordset
Private rsCD013 As New ADODB.Recordset
Private strPrgName As String  ''�{���W��
Private strPath As String     ''GICMIS1.INI���|
Private strDataName As String  ''�ɮצW��
Private strHeader As String   ''�O�����Y�ɪ���r
Private strChoose As String  ''
Private lngErrCount As Long   '' �O�����~����
Public blnSameFile As Boolean   '' �O�����P���d�O�O���O�P�@�ӯS���W�ٻP�N�X�A�O���ܫh�b�P�@�Ӥ�r��  append �i�h
Dim strBankId As String
Dim lngCount As Long       '' �O���Ҧ���r�ɿ�X���`����

Public intPara24 As Integer   ''  �O���O�_�ϥδC��h�b��B�z
Public strChooseMultiAcc As String   '' �o�ӰѼƥΨөӱ��z�� SO033 �һݪ����O���
Dim rsMedioBillnoNull As New Recordset   '' �o�ӰѼƥΨөӱ��z�� SO033�� mediobillno �Onull �����
Private blnUpdUcCode As Boolean
Private strUCCode As String
Private strUCName As String



Private Sub cboBillType_Click()
    If cboBillType.ListIndex = 2 Then
        cboSet.ListIndex = 0
        txtBatch.Enabled = True
        txtBatch.Text = "01"
    Else
        cboSet.Text = ""
        txtBatch.Enabled = False
        txtBatch.Text = ""
    End If

End Sub

Private Sub chkDuteDate_Click()
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
    On Error GoTo ChkErr
    Dim s As String
    Dim intTxtFileIndex  As Integer
    Dim lngTime As Long
        lngTime = Timer
        If Not IsDataOk Then Exit Sub
        
        '********************************************************************************************************************************************************
        '#3417 �q�l�ɶץX��,�n��J������](RefNo=4) By Kin 2007/12/05
        If Not GetRS(rsCD013, "Select * From " & TableOwnerName & "CD013 Where RefNo=4 And StopFlag<>1 Order By CodeNo Desc", cn, adUseClient, adOpenKeyset) Then Exit Sub
        If rsCD013.EOF Then
            blnUpdUcCode = False
        Else
            blnUpdUcCode = True
            strUCCode = rsCD013("CodeNo") & ""
            strUCName = rsCD013("Description") & ""
        End If
        '*********************************************************************************************************************************************************
        
        
        '' Call DefineRs   ''�إ߰O���� recordset ���c
        s = txtDataPath.Text
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(s) Then MsgBox "���| " & s & " ���s�b!", vbExclamation: Exit Sub
        '���~��Log
        If OpenFile(txtErrPath, 0, True) = False Then Exit Sub
        '��X�榡
        'objStorePara.uProcText = txtDataPath.Text
        objStorePara.uProcErrText = txtErrPath.Text
        objStorePara.uChkType = cboBillType.ListIndex
        If cboBillType.ListIndex = 0 Then
            If Not subCitiBankFlow Then file(0).Close: Exit Sub
        ElseIf cboBillType.ListIndex = 1 Then
            If Not subHSBCFlow Then file(0).Close: Exit Sub
        '���D��2763  �s�W��X�妸�榡
        ElseIf cboBillType.ListIndex = 2 Then
            If Not subCitiBankBatch Then file(0).Close: Exit Sub
        End If
        '' �N�����ұo����ƶ�J��r��
        file(0).Close
        Call ScrToRcd
        objStorePara.uUpdate = True
        
        If lngErrCount > 0 Then Call Shell("NotePad " & txtErrPath.Text, vbNormalFocus)
        Call msgResult(lngCount, lngErrCount, lngTime)      '��ܰ��浲�G
        Unload Me

    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Function subCitiBankFlow() As Boolean
    On Error GoTo ChkErr
    Dim rsCD037SpcNO  As New ADODB.Recordset  '' ���o SPCNO ���ߤ@�Ȫ��O���A�H�P��SPCNO���d�O�@�����X
    Dim rsCD037 As New ADODB.Recordset
    Dim s As String
    Dim intTxtFileIndex  As Integer
        ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' ���o SPCNO ���ߤ@�Ȫ��O���A�H�P��SPCNO���d�O�@�����X
        ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        '�H�Υd�����N�X,���X���b�S���N�X
        

        
        If Not GetRS(rsCD037SpcNO, "SELECT DISTINCT SPCNO FROM " & TableOwnerName & "CD037  WHERE SpcNO IS NOT NULL  ", cn, adUseClient) Then Exit Function
        If rsCD037SpcNO.EOF Or rsCD037SpcNO.BOF Then
            MsgBox "�H�Υd���������b�S���N�X���ݳ]�� "
            rsCD037SpcNO.Close
            Exit Function
        End If
        ' ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        '�H�Υd�����N�X,��X���b�S���W��<>NULL
        If Not GetRS(rsCD037, "SELECT * FROM " & TableOwnerName & "CD037 " & _
                "WHERE " & _
                "Spcno IS NOT NULL AND MerchantName IS NOT NULL " & _
                "ORDER BY CodeNO", cn, adUseClient) Then Exit Function
        intTxtFileIndex = 1
        ' ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        Screen.MousePointer = vbHourglass
        Dim lngSubCount As Long        '' �x�s�C�@�����P�����b�S���N�����B�z��Ƶ���
        Dim lngTotalAmount As Long   '' �x�s�C�@�����P�S���b�S���N���B�z��ƪ��B�`�X
        Dim blnNewRec As Boolean     '' �O���o�O�_���@���s��
        Do While Not rsCD037SpcNO.EOF
            rsCD037.Filter = ""
            rsCD037.Filter = "SPCNO = " & rsCD037SpcNO(0)
            lngSubCount = 0
            lngTotalAmount = 0
            blnNewRec = False
            Do While Not rsCD037.EOF
                If blnNewRec = False Then
                       Call DefineRs   ''�إ߰O���� recordset ���c
                End If
                 
                If Not BeginTran(rsCD037("Description"), rsCD037("SPCNO"), rsCD037("MerchantName"), intTxtFileIndex, lngSubCount, lngTotalAmount, blnNewRec) Then
                End If
                rsCD037.MoveNext
                blnNewRec = True
                DoEvents
            Loop
           
           '' �N�����ұo����ƶ�J��r��
           
            If lngSubCount > 0 Then
                rsCD037.MoveFirst
                ''       If OpenFile(txtDataPath, intTxtFileIndex, True, cboSet.ListIndex, rsCD037("MerchantName"), rsCD037("SPCNO")) = False Then
                
                If OpenFile(txtDataPath, intTxtFileIndex, True, cboSet.ListIndex, rsCD037("MerchantName"), rsCD037("CardID")) = False Then
                    MsgBox "�L�k�إߵn���S���N�X�� " & rsCD037("SPCNO") & " ����r�� !!", vbOKOnly + vbCritical, "�T�� "
                    Exit Function
                End If
                
                '' �H�U���o header �r��
                
                strHeader = rsCD037("MerchantName") & Space((9 - Len(rsCD037("MerchantName"))))
                strHeader = strHeader & Format(gdaSystemDate.Text, "yyyymmdd") & Format(lngSubCount, "000000")
                strHeader = strHeader & Format(lngTotalAmount, "0000000000") & cboSet.ListIndex + 1
                strHeader = strHeader & "autobilling" & Space(133) & "00"
                Call subWriteLine(intTxtFileIndex)   ''''' �N���o����Ƽg�J��r��
            End If
            rsCD037SpcNO.MoveNext
            intTxtFileIndex = intTxtFileIndex + 1
        Loop
        
        
        rsCD037.Close
        Set rsCD037 = Nothing
        rsCD037SpcNO.Close
        Set rsCD037SpcNO = Nothing
        subCitiBankFlow = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subCitiBankFlow")
End Function

Private Function subHSBCFlow() As Boolean
    On Error GoTo ChkErr
    Dim intTxtFileIndex As Integer
    Dim lngSubCount As Long
    Dim lngTotalAmount As Long
    Dim strFileName As String
        lngSubCount = 0
        lngTotalAmount = 0
        intTxtFileIndex = 1
        Call DefineRs2   ''�إ߰O���� recordset ���c
        If Not BeginTran("", txtSpcNO, txtMerchantName, intTxtFileIndex, lngSubCount, lngTotalAmount, True) Then
            'Exit Function
        End If
           
        '' �N�����ұo����ƶ�J��r��
        
        If lngSubCount > 0 Then
            strFileName = txtDataPath & "\" & txtMerchantName & "." & Format(Date, "yymmdd") & "." & Format(intTxtFileIndex, "00") & ".in"
            objStorePara.uProcText = strFileName
            Set file(intTxtFileIndex) = fso.CreateTextFile(strFileName, True)
            
            '' �H�U���o header �r��
            '1   �O���ѧO�X 1     'H'
            '2   �׺ݾ����X  8   ��J0 �� 9 �Ʀr(�e���W�ݥ����N��)
            '3   �ө��N��    15  ��J 0 �� 9�Ʀr�A �a�������ɪť�(�e���W�ө��N��)
            '4   ���O    1   T':TWD, 'U':USD(�w�]T)
            '5   �`����  6   ��J 000001-999999 �Ʀr
            '6   ���v���M����(A)����   6   ��J 000001-999999 �Ʀr(���b����1 . ���v�@�p��)
            '7   ���v���M����(A)���B   12  �p��2��A�a�k�����e���� 0(���b����1 . ���v�@�p��)
            '8   ���v��M����(S)����   6   ��J 000001-999999 �Ʀr(���b����3 . ���v�ýдڧ@�p��)
            '9   ���v��M����(S)���B   12  �p��2��A�a�k�����e���� 0(���b����3 . ���v�ýдڧ@�p��)
            '10  �ɵn�M����(O)���� 6   ��J 000001-999999 �Ʀr(���b����2 . �дڧ@�p��)
            '11  �ɵn�M����(O)���B 12  �p��2��A�a�k�����e���� 0(���b����2 . �дڧ@�p��)
            '12  �h�f�M����(R)���� 6   ��J 000001-999999 �Ʀr(���b����4 . �h�ڧ@�p��)
            '13  �h�f�M����(R)���B 12  �p��2��A�a�k�����e���� 0(���b����4 . �h�ڧ@�p��)
            '14  �O�d���    151 Space
            Dim lngSetCount(4) As Long, lngSetTotalAmt(4) As Long
            strHeader = "H" & GetString(txtMerchantName, 8) & GetString(txtSpcNO, 15)
            strHeader = strHeader & "T" & Format(lngSubCount, "000000")
            
            lngSetCount(cboSet.ListIndex + 1) = lngSubCount
            lngSetTotalAmt(cboSet.ListIndex + 1) = lngTotalAmount
            
            strHeader = strHeader & GetString(lngSetCount(1), 6, GIRIGHT, True) & GetString(lngSetTotalAmt(1), 10, GIRIGHT, True) & "00"
            strHeader = strHeader & GetString(lngSetCount(3), 6, GIRIGHT, True) & GetString(lngSetTotalAmt(3), 10, GIRIGHT, True) & "00"
            strHeader = strHeader & GetString(lngSetCount(2), 6, GIRIGHT, True) & GetString(lngSetTotalAmt(2), 10, GIRIGHT, True) & "00"
            strHeader = strHeader & GetString(lngSetCount(4), 6, GIRIGHT, True) & GetString(lngSetTotalAmt(4), 10, GIRIGHT, True) & "00"
            strHeader = strHeader & Space(151)
            Call subWriteLine(intTxtFileIndex)   ''''' �N���o����Ƽg�J��r��
        End If
        subHSBCFlow = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subHSBCFlow")
End Function

Private Sub Form_Load()
    On Error Resume Next
        blnSameFile = False
        lngErrCount = 0
        lngCount = 0
        cboBillType.ListIndex = 0
        Me.Caption = objStorePara.BankName & ""
        Call InitData
        RcdToScr
      ''���o�O�_�ϥΦh�C��
     
End Sub
Private Sub RcdToScr()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
      txtPayDate.Text = Format(Date, "dd")
        If Dir(strPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
            Set LogFile = fso.OpenTextFile(strPath & "\" & strPrgName & "Out.log", ForReading)
            With LogFile
            
                'If Not .AtEndOfStream Then txtPayDate.Text = .ReadLine & ""   '�дڤ��
                If Not .AtEndOfStream Then txtDiscountRate.Text = .ReadLine & ""   '�馩�v
                If Not .AtEndOfStream Then txtStatement.Text = .ReadLine & ""  '�Ȥ�b�椺�e
                '              If Not .AtEndOfStream Then txtInVoice.Text = .ReadLine & ""   '�Τ@�s��
                
                If Not .AtEndOfStream Then txtDataPath = .ReadLine & ""   '�����
                If Not .AtEndOfStream Then txtErrPath = .ReadLine & ""   '���D��
                If Not .AtEndOfStream Then txtSpcNO = .ReadLine & ""   '�ݥ����W��
                If Not .AtEndOfStream Then txtMerchantName = .ReadLine & ""   '�ө��W��
                  
            End With
            LogFile.Close
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Private Sub InitData()
  On Error GoTo ChkErr
'  Dim o As Object
'  Set o = CreateObject("CreditCardOut.clsInterface")
    With objStorePara
      Set cn = .Connection
      strBankId = .BankId
      'strBankName = .BankName
      'strCorpId = .CorpID
      gdaSystemDate.SetValue (Date)
      strChoose = .ChooseStr
      strPath = .ErrPath
   ''   txtSpcNO.Text = .uSpcNo
    End With
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
'        With cnd1
'                .FileName = txtDataPath.Text
'                .Filter = "��r��|*.txt"
'                .InitDir = strPath
'                .Action = 1
'                txtDataPath = .FileName
'        End With
    txtDataPath.Text = FolderDialog("����ɸ��| ")
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
        
        ''"""""""""""""""""""""""""""""""""
        .Fields.Append ("SequenceNo"), adBSTR, 6, adFldIsNullable
        .Fields.Append ("MerchantNo"), adBSTR, 9, adFldIsNullable
        .Fields.Append ("DiscountRate"), adBSTR, 5, adFldIsNullable
        .Fields.Append ("CardNo"), adBSTR, 16, adFldIsNullable
        .Fields.Append ("ExpiryDate"), adBSTR, 6, adFldIsNullable
        
        .Fields.Append ("SequenceNo2"), adBSTR, 15, adFldIsNullable
        .Fields.Append ("Filler1"), adBSTR, 8, adFldIsNullable
        .Fields.Append ("Amount"), adBSTR, 7, adFldIsNullable
        .Fields.Append ("StatementDescription"), adBSTR, 38, adFldIsNullable
        .Fields.Append ("Filler2"), adBSTR, 2, adFldIsNullable
        
        .Fields.Append ("MerchantDescription"), adBSTR, 20, adFldIsNullable
        .Fields.Append ("StatusCode"), adBSTR, 2, adFldIsNullable
        .Fields.Append ("ApprovalCode"), adBSTR, 6, adFldIsNullable
        .Fields.Append ("AuthorizationDate"), adBSTR, 8, adFldIsNullable
        .Fields.Append ("Filler3"), adBSTR, 16, adFldIsNullable
        
        .Fields.Append ("Filler4"), adBSTR, 6, adFldIsNullable
        .Fields.Append ("PaymentFlag"), adBSTR, 1, adFldIsNullable
        .Fields.Append ("ResultDescription"), adBSTR, 3, adFldIsNullable
        .Fields.Append ("BillingDate"), adBSTR, 2, adFldIsNullable
        .Fields.Append ("Filler5"), adBSTR, 2, adFldIsNullable
        
        .Fields.Append ("HeaderIndicator"), adBSTR, 2, adFldIsNullable
        .Open
    End With
    
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefineRs")
End Sub

Private Sub DefineRs2()
    On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        ''"""""""""""""""""""""""""""""""""
        .Fields.Append ("F01"), adVarChar, 1, adFldIsNullable
        .Fields.Append ("F02"), adVarChar, 6, adFldIsNullable
        .Fields.Append ("F03"), adVarChar, 8, adFldIsNullable
        .Fields.Append ("F04"), adVarChar, 15, adFldIsNullable
        .Fields.Append ("F05"), adVarChar, 6, adFldIsNullable
        
        .Fields.Append ("F06"), adVarChar, 1, adFldIsNullable
        .Fields.Append ("F07"), adVarChar, 19, adFldIsNullable
        .Fields.Append ("F08"), adVarChar, 4, adFldIsNullable
        .Fields.Append ("F09"), adVarChar, 3, adFldIsNullable
        .Fields.Append ("F10"), adVarChar, 12, adFldIsNullable
        
        .Fields.Append ("F11"), adVarChar, 6, adFldIsNullable
        .Fields.Append ("F12"), adVarChar, 2, adFldIsNullable
        .Fields.Append ("F13"), adVarChar, 20, adFldIsNullable
        .Fields.Append ("F14"), adVarChar, 20, adFldIsNullable
        .Fields.Append ("F15"), adVarChar, 40, adFldIsNullable
        
        .Fields.Append ("F16"), adVarChar, 1, adFldIsNullable
        .Fields.Append ("F17"), adVarChar, 10, adFldIsNullable
        .Fields.Append ("F18"), adVarChar, 3, adFldIsNullable
        .Fields.Append ("F19"), adVarChar, 9, adFldIsNullable
        .Fields.Append ("F20"), adVarChar, 9, adFldIsNullable
        
        .Fields.Append ("F21"), adVarChar, 7, adFldIsNullable
        .Fields.Append ("F22"), adVarChar, 16, adFldIsNullable
        .Fields.Append ("F23"), adVarChar, 36, adFldIsNullable
        .Open
        Dim intLoop As Integer
        Dim intMin As Long, intMax As Long
        intMin = 1
        For intLoop = 0 To .Fields.Count - 1
            intMax = intMin + .Fields(intLoop).DefinedSize - 1
            Debug.Print GetString((intLoop + 1), 2) & " " & GetString(.Fields(intLoop).DefinedSize, 3, GIRIGHT) & " " & GetString(intMin, 3, GIRIGHT) & "-" & GetString(intMax, 3, GIRIGHT)
            intMin = intMax + 1
        Next
    End With
        
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefineRs2")
End Sub

Private Function BeginTran(ByVal strCardName As String, ByVal strSpcNo As String, _
                                          ByVal strMerchantName As String, ByVal intIndexTxt As Integer, _
                                          ByRef pSubCount As Long, ByRef pTotalAmount As Long, ByVal pNewRec As Boolean) As Boolean

  On Error GoTo ChkErr
  
  Dim strsql As String
  Dim strOldBillNo As String
  Dim lngIndex As Long
  Dim lngSubCount As Long    '' �O���o�ӥd�O�O���O���ۦP���T����ƳQ��X�A���A�إߤ�r��
  
  Dim set34 As String   '' �p�G�O���v�P�д� �h�������� SHOULDAMT
  Dim strSequenceNumber As String   '' �O��  �h�C��渹
  Dim strSQLGroup   ''92/08/14   �վ�W�� �s�ժ��y�k�ھڬO�_�ҭ� intPara24 �@���� �ҥH�w�q�o���ܼ�
  Dim strMediaBillNONull  As String
  Dim lngCustIDErr As Long  '' ���o�Y����ƦbSO002A���Ƚs
  Dim rsCustIDErr As New ADODB.Recordset
  Dim rsSpcNo As New ADODB.Recordset
  Dim strSelectSPECNO As String
  ''2005/01/10  �[�J�P�_��  ���H�Υd���Ħ~��̤j��
  Dim strAccountID As String
  Dim lngCustID As Long
  Dim CountSameCustIDBillNO As Integer
  
  Dim theSingalAmount As Long  '' 2005/02/02 �Ψ��x�s�浧SO033��ơA
  
  Dim strCardNameSQL As String
  Dim lngerror As Long:    lngerror = 0
  
  Dim strWhereUCCode As String
  Dim strUpdUCCode As String
  Dim strStopYmWhere As String
  strAccountID = ""
  lngCustID = 0
  CountSameCustIDBillNO = 0
  BeginTran = False
    
'' 2005/02/02 ���ޥ��t���F  �@�߳��X��
   set34 = "  "
   
   ''' 2003/07/23 �s�W�Ȥ�s��(CustID)�H�δC��渹(MediaBillNo)�W��
   ''' 2003/08/14 �令SO002A ����ƨӷ��A�åB�NGROUP BY ���r�y�ζ}�A�����ϥδC��渹�P���ϥΪ�
   '''�ϥΥHMediaBillNO ���D�A�S���H Billno���D �A�������Ȥ�s����T �h�H  ���s���o��SO002A ���ǡACustStatusCode  �b SO002A �S�� �]���n�N SO002 ��i��
'' *****************************************************************************************************
    ''  2004/09/24 �[�JSO106.StopFlag = 0 ������� ����X�{�ⵧ
    ''  2004/09/24 �٬O�N��qmark���n�F�A������SO106�N�n�F
    ''  �H�U����qSQL �N�䤤�� A.ShouldAmt > 0  �����󮳱�
    
    strStopYmWhere = "MAX(DECODE(LENGTH(STOPYM),5,SUBSTR(STOPYM,2,4) || '0' || SUBSTR(STOPYM,1,1), " & _
                        "SUBSTR(STOPYM,3,4) || SUBSTR(STOPYM,1,2)))"
    
    If strCardName <> "" Then strCardNameSQL = " And C.CARDNAME  ='" & strCardName & "'"
    '#5055 ���M�S����,�����K�@�_��,�d�x�w���h���A�X�b By Kin 2009/04/30
    '#5218 �d�x�w���n��Nvl���覡 By Kin 2009/08/05
    '#5267 ���SO106�h���� BY Kin 2010/04/21
    If intPara24 = 0 Then

        '#5564 �W�[�ѦҸ�7,8�PCD013.PAYOK�N��w�� By Kin 2010/05/17
        '#5683 �W�[���RealStopDate �P PayKind By Kin 2010/08/06
        strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                " A.BillNO, sum(A.ShouldAmt)  ShouldAmt,  " & _
                " A.AccountNo,A.CUSTID, " & _
                strMinPayKind & " , " & strMinRealStopDate & _
                " From " & _
                TableOwnerName & "SO033 A ," & _
                TableOwnerName & "SO001 B ," & _
                TableOwnerName & "SO002 D,  " & _
                TableOwnerName & "CD013 " & _
                "Where " & strBankId & " AND " & _
                " A.CustID  = D.CUSTID AND " & _
                " A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                " A.UCCode Is Not Null  AND  " & _
                " A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) AND  " & _
                " NVL(CD013.PAYOK,0) = 0 AND " & _
                " A.SERVICETYPE = D.SERVICETYPE " & _
                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
                " Group by A.BillNO,A.AccountNo,A.CUSTID ORDER BY  A.BillNO DESC  "

        strsql = "SELECT A.BILLNO,A.SHOULDAMT,A.ACCOUNTNO," & _
            " Min(A.PayKind) PayKind,Min(A.RealStopDate) RealStopDate, " & _
            strStopYmWhere & " CardExpDate " & _
            " FROM (" & strsql & ") A," & TableOwnerName & "SO106 C " & _
            " WHERE A.ACCOUNTNO=C.ACCOUNTID AND C.STOPFLAG <> 1 " & _
            " AND SnactionDate IS NOT NULL AND A.CUSTID=C.CUSTID " & strCardNameSQL & _
            " GROUP BY A.BILLNO,A.SHOULDAMT,A.ACCOUNTNO " & _
            " ORDER BY A.BILLNO DESC "
    Else
     'So033:�����������,So001:�Ȥ�����,So002A:�Ȥ�b����,So002:�Ȥ�A�Ȱ򥻸��,So106:�H�Υd����b�ӽаO��
     '��X�Ȥ᥿�������`���B
        '#5267 ���SO106�h���� By Kin 2010/04/21
        '#5564 �W�[�ѦҸ�7,8�PCD013.PAYOK�N��w�� By Kin 2010/05/17
        strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                "sum(A.ShouldAmt)  ShouldAmt,  " & _
                "A.AccountNo,A.MediaBillNo,A.CustId, " & _
                strMinPayKind & " , " & strMinRealStopDate & _
                " From " & _
                TableOwnerName & "SO033 A ," & _
                TableOwnerName & "SO001 B ," & _
                TableOwnerName & "SO002 D,  " & _
                TableOwnerName & "CD013 " & _
                " Where " & strBankId & " AND " & _
                " A.CustID  = D.CUSTID AND " & _
                " A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                " A.UCCode Is Not Null AND  " & _
                " A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) And " & _
                " NVL(CD013.PAYOK,0) = 0 AND " & _
                " A.SERVICETYPE = D.SERVICETYPE  " & _
                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
                " Group by A.AccountNO,A.MediaBillNo,A.CUSTID "
                
        strsql = "SELECT A.SHOULDAMT,A.ACCOUNTNO," & _
                    " A.MEDIABILLNO," & strStopYmWhere & " CardExpDate, " & _
                    " Min(PayKind) PayKind,Min(RealStopDate) RealStopDate " & _
            " FROM (" & strsql & ") A," & TableOwnerName & "SO106 C " & _
            " WHERE A.ACCOUNTNO=C.ACCOUNTID AND C.STOPFLAG <> 1 " & _
            " AND SnactionDate IS NOT NULL AND A.CUSTID=C.CUSTID " & strCardNameSQL & _
            " GROUP BY A.SHOULDAMT,A.ACCOUNTNO,A.MEDIABILLNO " & _
            " ORDER BY A.MEDIABILLNO DESC "
            
    End If
    
    '**************************************************************************************************************
    '#3417 ��sUCCode�PUCName�򥻪� Where���� By Kin 2007/12/05
    strWhereUCCode = strBankId & strCardNameSQL & " AND " & _
                    "A.CustID  = C.CUSTID AND " & _
                    "A.CustID  = D.CUSTID AND " & _
                    "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                    "A.UCCode Is Not Null AND  " & _
                    "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountID  AND C.StopFlag = 0  " & _
                    IIf(Len(strChoose) = 0, "", " And ") & strChoose
    '***************************************************************************************************************

 ''  2003/08/14 �令 �A�N mediabillno   ��ƨ��X �o�سo�@���h���@�� BillNO �M���J��mediabillno��null �Ȫ��ɭԡA��J�C��渹
    '#5055 �S����,�����K�@�_��,�n�L�o���d�x�w�� By Kin 2009/04/30
    '#5218 �d�x�w���n��Nvl�覡 By Kin 2009/08/05
    If intPara24 = 1 Then
    
                        ''�o�@�q�������e�z���D SQL (strsql)  ���� group  �A�[�W Billno
        '#5564 �W�[�ѦҸ�7,8�PCD013.PAYOK�N��w�� By Kin 2010/05/17
        strMediaBillNONull = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                             "C.AccountID,A.MediaBillNo,A.BillNO,MAX(A.MEDIABILLNO)   " & _
                             "From " & _
                              TableOwnerName & "SO033 A ," & _
                              TableOwnerName & "SO001 B ," & _
                              TableOwnerName & "SO106 C,  " & _
                              TableOwnerName & "SO002 D,  " & _
                              TableOwnerName & "CD013 " & _
                             "Where " & strBankId & strCardNameSQL & " AND " & _
                             "A.CustID  = C.CUSTID AND " & _
                             "A.CustID  = D.CUSTID AND " & _
                             "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                             "A.UCCode Is Not Null  AND " & _
                             "A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) And " & _
                             "NVL(CD013.PAYOK,0) = 0 AND " & _
                             "A.SERVICETYPE=D.SERVICETYPE  AND A.AccountNo = C.AccountID  " & _
                             IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & "  AND A.MEDIABILLNO IS NULL" & _
                             " GROUP BY C.ACCOUNTID,A.MEDIABILLNO,A.BILLNO "
                                            
                                            
        strChooseMultiAcc = " A.CancelFlag = 0 AND " & _
                            " A.UCCode Is Not Null  " & _
                            IIf(Len(strChooseMultiAcc) = 0, "", " And ") & strChooseMultiAcc
                
        If GetRS(rsMedioBillnoNull, strMediaBillNONull, cn) Then
            Do While Not rsMedioBillnoNull.EOF
                If IsNull(rsMedioBillnoNull("MediaBillNo")) Then
                     strSequenceNumber = GetInvoiceNo2("SO033_MediaBillNo", cn)
                     cn.Execute ( _
                                    "UPDATE " & TableOwnerName & "SO033 A SET " & _
                                    "A.MediaBillNo  ='" & strSequenceNumber & "'  WHERE   " & _
                                    "A.BillNo = '" & rsMedioBillnoNull("BillNO") & "'   AND  " & strChooseMultiAcc & " AND " & _
                                    "A.AccountNO ='" & rsMedioBillnoNull("AccountID") & "' ")
                    rsMedioBillnoNull.MoveNext
                End If
            Loop
                     
        End If
        rsMedioBillnoNull.Close
        Set rsMedioBillnoNull = Nothing
    End If
    
    If intPara24 = 0 Then
        strChooseMultiAcc = " A.CancelFlag = 0 AND " & _
                            "A.UCCode Is Not Null   " & _
                             IIf(Len(strChooseMultiAcc) = 0, "", " And ") & strChooseMultiAcc
    End If
    
    Set rsTmp = cn.Execute(strsql)
    
     ''  2003/07/23    �o�@�q���o�u�� SO033 ���z����  ........
    
    If rsTmp.BOF Or rsTmp.EOF Then
        blnNodata = True
        objStorePara.uNoData = True
    Else
        blnNodata = False
        Dim strSN As String  '' ���w�ǦC��
        Dim lngTotal As Long '' �֭p���Y�� Total Amount ����`�X����B
        Dim lngDoubleMediaNO As Long '' �o�ӭȰO�����ХX�檺��l���

        lngDoubleMediaNO = 0
        If pNewRec = False Then
             lngTotal = 0
             lngSubCount = 0
        Else
            lngTotal = pTotalAmount
            lngSubCount = pSubCount
        End If
                                  
        Do While Not rsTmp.EOF
            Dim strCardExpDate As String
            Dim iCED As Integer
            Dim dteCED As Date
            Dim strErrMessage As String
            Dim strBillNOField As String
            
            If intPara24 = 0 Then
                 strBillNOField = "BillNO"
            Else
                 strBillNOField = "MediaBillNO"
            End If
                 
           ''  20050110 �[�J�_�� ���P�b���H�Υd�ح�����~��̤j�� �ѩ�SQL�w�gDESC�Ƨ�
           ''  �]���p�G�渹�O�ۦP��, ��ܨ��X���O�ۦP���ⵧ�A�N���n�~�򩹤U�F�A
           '' ���챵�U�Ӫ� �@���O���@�˪��A�X��
            '' ************************************************************************
            Dim strSopYMMaxString As String: lngerror = 1
            If strAccountID = rsTmp(strBillNOField) Then
                lngerror = 2
                GoTo NextLoop
            Else
                lngerror = 3
                strAccountID = rsTmp(strBillNOField)
                lngerror = 4
                '' 20050205 �վ�  ���Ĥ@���̤j�� �]����Ӫ�����άO��r���ƥX�ӷ|�X���D
                
                ''   strSopYMMaxString = GetStopYMDateString(rsTmp(strBillNOField), strCardName, set34)
                '#5267 �w�g���Max(SO106.StopYm),�ҥH���ݭn���q�F By Kin 2010/04/21
                'strSopYMMaxString = GetStopYMDateString(strAccountID, strCardName, set34)
                strSopYMMaxString = rsTmp("CardExpDate") & ""
                lngerror = 45
            End If
            '' ************************************************************************
      '     ''2005/01/26�N��z�n�����  �̷ӳ�ڱƧǤ���  �N�P�渹���t���X��
           '''Call LogSameNOPMAmount(rsTmp(strBillNOField), strCardName, set34)
           ''2005/02/02 �N��z�n�����  �̷ӳ�ڱƧǤ���  �N�P�渹���t���X��
           
           lngerror = 5
           '���D��2868 �q�XLog�ɡA�Ƚs�p�G���h�b��|�P�_���~�A�h�W�[�HSO033���Ƚs�P�_ 2006/11/15 for Jacy
            Call GetRS(rsCustIDErr, "SELECT  distinct SO001.CUSTNAME , SO106.CUSTID FROM " & TableOwnerName & "SO106," & TableOwnerName & "SO001  " & _
                                 "WHERE " & _
                                 "SO106.AccountID ='" & rsTmp("AccountNO") & "'  AND  SO106.CUSTID = SO001.CUSTID And SO106.CUSTID IN (" & _
                                 "Select DISTINCT CUSTID From " & TableOwnerName & "SO033 Where " & strBillNOField & "='" & _
                                  rsTmp(strBillNOField) & "' And AccountNo='" & rsTmp("AccountNO") & "')", cn)
            If rsCustIDErr.RecordCount <= 0 Then
                Call GetRS(rsCustIDErr, "Select SO001.CUSTNAME,SO033.CUSTID FROM " & TableOwnerName & "SO001," & TableOwnerName & "SO033 " & _
                                       "Where SO001.CUSTID=SO033.CUSTID And SO033.AccountNo='" & rsTmp("AccountNO") & "' And " & strBillNOField & "='" & _
                                       rsTmp(strBillNOField) & "'", cn)
                                       
                lngErrCount = lngErrCount + 1: lngerror = 6.5
                strErrMessage = "�Ȥ�b������ :�渹  " & rsTmp(strBillNOField) & " �Ȥ�m�W " & rsCustIDErr("CustName") & " �Ȥ�s��  " & rsCustIDErr("CustID")
                file(0).WriteLine strErrMessage: lngerror = 6.8
                GoTo NextLoop
            End If
            '#5683 �p�G�����I���p��e�����󤣭n�X�b By Kin 2010/08/06
            If Not IsPayKindOK(rsTmp, gdaSystemDate.GetValue & "", giAll, intPara24) Then
                lngErrCount = lngErrCount + 1
                strErrMessage = "�����I�����~ :�渹  " & rsTmp(strBillNOField) & " �Ȥ�m�W " & rsCustIDErr("CustName") & _
                        " �Ȥ�s��  " & rsCustIDErr("CustID") & " �����I��� : " & rsTmp("RealStopDate") & _
                        " ú�I���O : �{�I�� "
                file(0).WriteLine strErrMessage: lngerror = 6.9
                GoTo NextLoop
            End If
            
   lngerror = 6
            If IsNull(rsTmp("AccountNO")) Then
                lngErrCount = lngErrCount + 1:    lngerror = 7
                strErrMessage = "�H�Υd�d���ť� : �渹�@" & rsTmp(strBillNOField) & " �Ȥ�m�W " & rsCustIDErr("custname") & " �Ȥ�s��  " & rsCustIDErr("CustID")
                file(0).WriteLine strErrMessage:     lngerror = 8
                GoTo NextLoop
            End If
           
           
           ''20050205 �o�@��O���O�]�n�ﱼ �O   ************************************
            '***********************************
           ' ************************************
         ''  strCardExpDate = GetNullString(rsTmp("CardExpDate"), giStringV, giAccessDb)
                      '' 20050205�[�J�o�@��令��GetStopYMDateString���o���̤j��
                      
            strCardExpDate = strSopYMMaxString
            If strCardExpDate = "NULL" Then
                lngErrCount = lngErrCount + 1
                strErrMessage = "�H�Υd��Ƥ����T : �渹�@" & rsTmp(strBillNOField) & " �Ȥ�m�W " & rsCustIDErr("custname") & " �Ȥ�s��  " & rsCustIDErr("CustID")
                file(0).WriteLine strErrMessage
                GoTo NextLoop
            End If
            lngerror = 9
            '#5267 ���X�ɤw�g�榡�n�F,���ݭn�A��F By Kin 2010/04/21
            'strCardExpDate = Right(strSopYMMaxString, 4) & "/" & Replace(strSopYMMaxString, Right(strSopYMMaxString, 4), "")
            strCardExpDate = Left(strSopYMMaxString, 4) & "/" & Right(strSopYMMaxString, 2)
            lngerror = 10
           
            dteCED = CDate(strCardExpDate)
            If dteCED < CDate(Format(gdaSystemDate.Text, "YYYY/MM")) Then
                lngErrCount = lngErrCount + 1
                strErrMessage = "�H�Υd�L�� : �渹�@" & rsTmp.Fields(strBillNOField) & " �Ȥ�m�W�@" & rsCustIDErr("custname") & " �H�Υd�� " & rsTmp.Fields("AccountNO") & " �����@" & Format(rsTmp.Fields("CardExpDate"), "000000") & " �Ȥ�s��  " & rsCustIDErr("CustID")
                file(0).WriteLine strErrMessage
                If chkDuteDate.Value = 0 Then GoTo NextLoop
            End If
          
'           strCardExpDate = rsTmp("CardExpDate")
'           iCED = IIf(Len(strCardExpDate) = 5, 1, 2)
'           strCardExpDate = Right(strCardExpDate, 4) & "/" & Mid(strCardExpDate, 1, iCED)
          
          '' 2005/02/02 �p�G�[�`���B�p��s �@LOG���X��
          lngerror = 11
          Dim strtest As String
          strtest = rsTmp(strBillNOField)
          lngerror = 111
          theSingalAmount = GetSumAmout(rsTmp(strBillNOField), strCardName, set34)
          lngerror = 112
          If theSingalAmount < 0 Or theSingalAmount = 0 Then
                    strErrMessage = " �t���άO���B����s�G�渹�@" & rsTmp(strBillNOField)
                               lngerror = 12

                    lngErrCount = lngErrCount + 1
                    file(0).WriteLine strErrMessage
                               lngerror = 12

                    GoTo NextLoop
          End If
            lngerror = 13
        
        '���D�� 2852 �ˬd�S�����O�_���T�p�G�����T�A���\���Ʒ|�X��(���q���Ӧb�U��,�{�b����o��ӳB�z) For Jacy 2006/11/08
        If cboBillType.ListIndex = 2 Then
            strSpcNo = "Select SPCNO From CD037 WHERE CARDID IN (Select DISTINCT CARDCODE FROM " _
                       & TableOwnerName & " So106 A," & _
                        TableOwnerName & " So002A B Where A.ACCOUNTID='" & rsTmp("AccountNO") & "'" & _
                        " AND A.CUSTID=B.CUSTID AND A.ACCOUNTID=B.ACCOUNTNO AND A.CUSTID=" & rsCustIDErr("CustID") & " ) "
            If Not GetRS(rsSpcNo, strSpcNo, cn) Then Exit Do
            If rsSpcNo.RecordCount <= 0 Then
                strErrMessage = " ���b�S���N�X�����T�G�渹�@" & rsTmp(strBillNOField) & " �Ȥ�m�W�@" & rsCustIDErr("custname") & "  �Ȥ�s��  " & rsCustIDErr("CustID")
                lngErrCount = lngErrCount + 1
                file(0).WriteLine strErrMessage
                GoTo NextLoop
            Else
                strSpcNo = Trim(GetString(rsSpcNo(0), 15))
            End If
        End If
        lngCount = lngCount + 1
        lngSubCount = lngSubCount + 1
        If cboBillType.ListIndex = 0 Then
            With rsDefTmp
                 .AddNew
                 .Fields("SequenceNo") = Format(lngSubCount, "000000")
                 '' SequenceNo(1-6)
                 '#3201 �]��CD037.SPCNO�����ܬ�15�|�v�Q��X��榡�A�ҥH��榡��SPCNO�ɡA�h�W�[Mid��ơA�j�����e���E�X By Kin 2007/08/03
                 .Fields("MerchantNo") = Mid(Format(strSpcNo, "000000000"), 1, 9)
                 ''MerchantNo(7-15)
                 .Fields("DiscountRate") = Format(txtDiscountRate.Text, "00000")
                 ''DiscountRate(16-20)
                 .Fields("CardNo") = IIf(Len(rsTmp("AccountNO")) > 0, Format(rsTmp("AccountNO"), "0000000000000000"), "0000000000000000")
                 ''CardNo(21-36)
                 '' 20050205 MARK�H�W�����@��A�令�H�U��
                 strCardExpDate = Replace(strCardExpDate, "/", "")
                 If Len(strCardExpDate) = 5 Then
                     strCardExpDate = Left(strCardExpDate, 4) & "0" & Right(strCardExpDate, 1)
                 End If
                 .Fields("ExpiryDate") = Right(strCardExpDate, 2) & Left(strCardExpDate, 4)
                 ''ExpiryDate(37-42)
                 .Fields("SequenceNo2") = Format(lngSubCount, "000000000000000")
                 ''SequenceNo2(43-57)
                 .Fields("Filler1") = Space(8)
                 '' Filler(58-65)
                 ''2005/02/02 �H�U���B���ȧ令����SO033�����B �קK�{�����o���Ъ��B  !!
                 .Fields("Amount") = Format(theSingalAmount, "0000000")
                 lngTotal = lngTotal + theSingalAmount
                 '' Amount(66-72)
                  Dim strSD As String
                  Dim intlength As Integer
                  
                  '""�ഫ�������T���ת����Φr��'''''''''''''''''''''''''
                  strSD = StrConv(txtStatement.Text, vbFromUnicode)
                  intlength = LenB(strSD)
                  strSD = txtStatement.Text & Space(38)
                   ''""""""""""""""""""""""""""""""""""""""""""""""""""""
                 .Fields("StatementDescription") = StrConv(LeftB(StrConv(StrConv(strSD, vbWide), vbFromUnicode), 38), vbUnicode)
                 
                 .Fields("Filler2") = Space(2)
                 
                '' 2003/07/23 �H�U�i��渹��Ȫ��ާ@�A�P�_�O�_�H�C��渹���\��J��
                '' �p�G�O�h�H�C��渹�J��A�_�h�ഫ���³渹����A�J
                '' 2003/08/14 �H�U�����e�A���o�渹���e�A�p�G�ҥδC�骺�\��A�N������MediaBillNO
                '���D��2861 �p�G�O�ϥδC�^�渹�A�ץX�ɷ|�C�^�渹�|�q���X�ӡA�]��Else�U������A���p�߳Q�����F
                 If intPara24 = 0 Then
                      Dim strBillNOOld As String
                      strBillNOOld = Trim(CStr(Val(Left(rsTmp("BillNO"), 6)) - 191100)) & _
                                     Mid(rsTmp("BillNO"), 7, 1) & Format(Right(rsTmp("BillNO"), 6), "000000")
                     .Fields("MerchantDescription") = strBillNOOld & Space(20 - Len(strBillNOOld))
                 Else
                     .Fields("MerchantDescription") = GetString(rsTmp("MediaBillNo"), 20, giLeft, False)
                 End If
                 ''MerchantDescriptione(113-132)
                 .Fields("StatusCode") = Space(2)
                 ''StatusCode(133-134)
                 .Fields("ApprovalCode") = Space(6)
                 ''ApprovalCode(135-140)
                 .Fields("AuthorizationDate") = Space(8)
                  ''AuthorizationDate(141-148)
                 .Fields("Filler3") = Space(16)
                  ''Filler(149-164)
                 .Fields("Filler4") = Space(6)
                  ''Filler(165-170)
                 .Fields("PaymentFlag") = Space(1)
                 .Fields("ResultDescription") = Space(3)
                 .Fields("BillingDate") = Format(txtPayDate.Text, "00")
                 .Fields("Filler5") = Space(2)
                 .Fields("HeaderIndicator") = "01"
                 .Update
            End With
         Else
            'HSB�榡
            If cboBillType.ListIndex = 1 Then
                If Not BeginTran2(lngSubCount, lngTotal, strCardExpDate, theSingalAmount) Then Exit Do
            Else
                If Not BeginTran3(lngSubCount, lngTotal, strCardExpDate, theSingalAmount, strSpcNo) Then Exit Do
                rsSpcNo.Close
                Set rsSpcNo = Nothing
            End If
         End If
         
'         TableOwnerName & "SO033 A ," & _
'                TableOwnerName & "SO001 B ," & _
'                TableOwnerName & "SO002A C,  " & _
'                TableOwnerName & "SO002 D,  " & _
'                TableOwnerName & "SO106 E  " & _

        '**********************************************************************************************************************************
        '#3417 �p�GCD013.REFNO���]�t��4,�h��sUCCode�PUCName By Kin 2007/12/06
        If blnUpdUcCode Then
           '#3892 ��s�t���ܺC,�վ�y�k,�쥻��Exists,�{�b���= By Kin 2008/06/03
           '#4388 �W�[��s���ʤH���P���ʮɶ� By Kin 2009/04/30
           '#5218 �d�x�w���n��Nvl���覡 By Kin 2009/08/05
           '#5564 �W�[�ѦҸ�7,8�PCD013.PAYOK�N��w�� By Kin 2010/05/17
            strUpdUCCode = "UPDATE " & TableOwnerName & "SO033 Set UCCode=" & strUCCode & ",UCName='" & strUCName & "'" & _
                           ",UPDEN='" & objStorePara.uUPDEN & "',UPDTime='" & objStorePara.uUPDTime & "' " & _
                        " Where " & strBillNOField & " In (Select A." & strBillNOField & " From " & _
                                        TableOwnerName & "SO001 B ," & _
                                        TableOwnerName & "SO106 C," & _
                                        TableOwnerName & "SO002 D," & _
                                        TableOwnerName & "SO033 A," & _
                                        TableOwnerName & "CD013 " & _
                                        " Where " & strWhereUCCode & _
                                        " And A." & strBillNOField & "='" & rsTmp(strBillNOField) & "'" & _
                                        " And C.AccountID='" & rsTmp("AccountNO") & "'" & _
                                        " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) " & _
                                        " AND NVL(CD013.PAYOK,0) =0 " & _
                                        IIf(rsTmp("CardExpDate") & "" <> "", _
                                        " And Decode(Length(C.StopYM),5,Substr(C.StopYM,2,4)||'0'||Substr(C.StopYM,1,1),Substr(C.StopYM,3,4)||Substr(C.StopYM,1,2))=" & rsTmp("CardExpDate"), "") & ")"
                                    
            cn.Execute strUpdUCCode
        End If
        '**********************************************************************************************************************************
NextLoop:
        rsTmp.MoveNext
        Loop
        pSubCount = lngSubCount
        pTotalAmount = lngTotal
        
        BeginTran = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
  End If
  
Exit Function
ChkErr:
    Call CloseFS(intIndexTxt)
    Call ErrSub(Me.Name, "BeginTran" & lngerror & "-" & strAccountID)
End Function

Private Function BeginTran2(ByRef lngSubCount As Long, ByRef lngTotal As Long, _
    ByVal strCardExpDate As String, theSingalAmount As Long) As Boolean
    On Error GoTo ChkErr
    
        With rsDefTmp
            .AddNew
            '1   �O���ѧO�X 1     'D'
            .Fields("F01") = "D"
            '2   ����Ǹ�    6   ��J 0 �� 9�Ʀr�A�a�k�����e���� 0(�y����)
            .Fields("F02") = Format(lngSubCount, "000000")
            '3   �׺ݾ����X  8   ��J0 �� 9 �Ʀr�r(�e���W�ݥ����N��)
            .Fields("F03") = GetString(txtMerchantName, 8)
            '4   �ө��N��    15  ��J 0 �� 9�Ʀr�A �a�������ɪť�(�e���W�ө��N��)
            .Fields("F04") = GetString(txtSpcNO, 15)
            '5   �����  6   YYMMDD(�褸�~) (�e���W�t�Φ~��+�дڤ�)
            .Fields("F05") = Format(Date, "YYMM") & txtPayDate
            '6   ���v�M��覡    1   A:�u���v���M��(1 . ���v), S:���v��M�� (3 . ���v�ýд�),O:�ɵn�M��(2 . �д�),  R:�h�f�M��(4 . �h��)(�e�����b�����ﶵ�M�w)
            Select Case cboSet.ListIndex
                Case 0          '1 .���v
                    .Fields("F06") = "A"
                Case 1          '2 .�д�
                    .Fields("F06") = "O"
                Case 2          '3 .���v�ýд�
                    .Fields("F06") = "S"
                Case 3          '4 .�h��
                    .Fields("F06") = "R"
            End Select
            '7   �d    ��    19  �a��,�����k�ɪť�
            .Fields("F07") = GetString(rsTmp("AccountNo"), 19)
            '8   ���Ĵ���    4   MMYY, MBF(Must be full)(�褸�~)
            strCardExpDate = Format(strCardExpDate, "MMYY")
            .Fields("F08") = Left(strCardExpDate, 2) & Right(strCardExpDate, 2)
            '9   CVC2    3   ��J CVC2 or Space(�w�]Spac)
            .Fields("F09") = Space(3)
            '10  ��    �B    12  �a�k���ɳ̫��X�p�Ʀ��
            .Fields("F10") = GetString(theSingalAmount, 10, GIRIGHT, True) & "00"
            lngTotal = lngTotal + theSingalAmount
            '11  �֭㸹�X    6   �����v��J'999999'or �w���o���v�X�h������J���v�X(�����v��дڥ�)
            .Fields("F11") = "999999"
            '12  �^ �_ �X    2   �����v��J'99'or �w���v��J'00'(�J�b�{���O�O��00�ɤ~�i�H�J��)
            .Fields("F12") = "99"
             '13  ���d�H�b��C�L���  20  ������ƥi�L�b���d�H��b��W(�e���W�b�椺�e�L���кI��)
            .Fields("F13") = GetString(txtStatement, 20)
            '14  �ө��O�d���    20  �ө��ۦ�B��(Space)(�C��渹�ȪA�t�ΤJ�b�̾�)
            If intPara24 = 0 Then
                .Fields("F14") = GetString(rsTmp("BillNo"), 20)
            Else
                .Fields("F14") = GetString(rsTmp("MediaBillNo"), 20)
            End If
            '15  �H�u���v���    40  Context Length  Note
            '            Name    8   �a�������ɪť�
            '            ID  10  �a�������ɪť�
            '            DOB 7   YYYMMDD(����~),�a�k�������ɪť�
            '            Tel 15  �a�������ɪť�
            '            �����b�G���H�u���v�ɨϥ� (�ثe�t�εL�ϥζ�ť�)
            .Fields("F15") = Space(40)
            '16  �������O    1   1:������� , 0:�D�������(�ثe�t�εL�ϥΤ�����0)
            .Fields("F16") = "0"
            '17  ���~�N�X    10  �a�������� '0'(�ثe�t�εL�ϥζ�ť�)
            .Fields("F17") = Space(10)
            '18  ����    3   �����I�ڥ�, Right justified, Start with zeros(�D���^��) (�ثe�t�εL�ϥζ�ť�)
            .Fields("F18") = Space(3)
            '19  �C�����B    9   �����I�ڥ�, Right justified, Start with zeros(�D���^��) (�ثe�t�εL�ϥζ�ť�)
            .Fields("F19") = Space(9)
            '20  �Ĥ@�����B  9   �����I�ڥ�, Right justified, Start with zeros(�D���^��) (�ثe�t�εL�ϥζ�ť�)
            .Fields("F20") = Space(9)
            '21  ����O�v    7   �����I�ڥ�, Right justified, Start with zeros(�D���^��) (�ثe�t�εL�ϥζ�ť�)
            .Fields("F21") = Space(7)
            '22  �����X(TC)  16  ��������дڥ�(�ثe�t�εL�ϥζ�ť�)
            .Fields("F22") = Space(16)
            '23  �O�d���    36  Space(�ثe�t�εL�ϥζ�ť�)
            .Fields("F23") = Space(36)
            '24  CR+LF   2   Chr(13)+Chr(10)
            '.Fields("F24") = ""
        End With
        BeginTran2 = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "BeginTran2"
End Function

Private Function subWriteLine(ByVal intIndexTxt As Integer) As Boolean

  On Error GoTo ChkErr
  
    Dim varData As Variant
    Dim strData As String
    Dim lngLoop As Long
    Dim lngLoopi As Long
    Dim strContent As String  '' �C�@���g�J�����

    subWriteLine = False
    
    With rsDefTmp
        
         If .BOF Or .EOF Then Exit Function
         .MoveFirst
         If blnSameFile = False Then
              file(intIndexTxt).WriteLine strHeader     ''���r�ɶ�J����
         End If
         For lngLoop = 0 To .RecordCount - 1
            strContent = ""
            For lngLoopi = 0 To rsDefTmp.Fields.Count - 1
                strContent = strContent & rsDefTmp(lngLoopi)
            Next
            file(intIndexTxt).WriteLine strContent
            rsDefTmp.MoveNext
            DoEvents
         Next
    End With
    subWriteLine = True
    On Error Resume Next
    rsDefTmp.Close
    Set rsDefTmp = Nothing
    file(intIndexTxt).Close
    Exit Function
ChkErr:
    subWriteLine = False
    Call ErrSub(Me.Name, "subWriteLine")
End Function

Private Sub ScrToRcd()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
        
        Set LogFile = fso.CreateTextFile(strPath & "\" & strPrgName & "Out.log", True)
        With LogFile
                '.WriteLine (txtPayDate.Text)      '�дڤ��
                .WriteLine (txtDiscountRate.Text) '�馩�v
                .WriteLine (txtStatement.Text)    '�Ȥ�b�椺�e
                '.WriteLine (txtInVoice.Text)     '�Τ@�s��
                .WriteLine (txtDataPath)         '����ɸ��|
                .WriteLine (txtErrPath.Text)     ' ���D�Ѧ��ɸ��|
                .WriteLine (txtSpcNO.Text)  '�ݥ����W��
                .WriteLine (txtMerchantName.Text) '�ݥ����W��
        End With
        LogFile.Close
        Set LogFile = Nothing
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Function IsDataOk() As Boolean
    Dim strErrMsg As String
        '�p�G���O��ܪ�X�妸�A�h���b����1.���v�P2.�дڵL�k�ϥ�
        If cboBillType.ListIndex <> 2 Then
            If cboSet.ListIndex = 0 Or cboSet.ListIndex = 1 Then
                MsgBox "���b���� (1.���v 2.�д�) �ثe�L�k�ϥ� !!", vbOKOnly + vbInformation, "�T���@"
                cboSet.SetFocus
                Exit Function
            End If
        End If
        If Len(cboSet.Text) = 0 Then
            MsgBox "�п�ܤ@�ص��b���� !!", vbOKOnly + vbInformation, "�T�� "
            cboSet.SetFocus
            Exit Function
        End If
        
        If Not IsNumeric(txtPayDate.Text) Then
            MsgBox " �дڤ���榡�����T�@!! ", vbOKOnly + vbInformation, "�T���@"
            txtPayDate.SetFocus
            Exit Function
        End If
        
        If Not IsDate(Format(Date, "yyyy/mm") & "/" & txtPayDate.Text) Then
            MsgBox " �дڤ���榡�����T�@!! ", vbOKOnly + vbInformation, "�T���@"
            txtPayDate.SetFocus
            Exit Function
        End If
        
        If Len(gdaSystemDate.Text) = 0 Then
            MsgBox " �t�Τ�����ݦ��ȡ@!! ", vbOKOnly + vbInformation, "�T���@"
            gdaSystemDate.SetFocus
            Exit Function
        End If
         
        If Not IsNumeric(txtDiscountRate.Text) Then
            MsgBox " �馩�v���ݬ��Ʀr�@!! ", vbOKOnly + vbInformation, "�T���@"
            txtDiscountRate.SetFocus
            Exit Function
        End If
  
        If cboBillType.ListIndex <> 0 Then
            If Len(txtSpcNO.Text & "") = 0 Then strErrMsg = "�ө��N��": txtSpcNO.SetFocus: GoTo Warning
            If Len(txtMerchantName.Text & "") = 0 Then strErrMsg = "�ݥ����N��": txtMerchantName.SetFocus: GoTo Warning
        End If
        '�p�G��ܪ�X�妸�A�h���b�����u�����v
        If cboBillType.ListIndex = 2 Then
            If cboSet.ListIndex <> 0 Then
                MsgBox "���b�����u���� 1.���v", vbOKOnly + vbInformation, "�T��"
                cboSet.ListIndex = 0
                cboSet.SetFocus
                Exit Function
            End If
            If Trim(txtBatch) = "" Then
                strErrMsg = "�帹"
                txtBatch.SetFocus
                GoTo Warning
            End If
        End If
        If Len(cboSet.Text) = 0 Then strErrMsg = "���b����": cboSet.SetFocus: GoTo Warning
        If Len(txtPayDate.Text) = 0 Then strErrMsg = "�дڤ��": txtPayDate.SetFocus: GoTo Warning
        If Len(txtDiscountRate.Text) = 0 Then strErrMsg = "�馩�v": txtDiscountRate.SetFocus: GoTo Warning
        If Len(txtDataPath.Text) = 0 Then strErrMsg = "����ɸ��|": txtDataPath.SetFocus: GoTo Warning
        If Len(txtErrPath.Text) = 0 Then strErrMsg = "���D�Ѧ��ɸ��|": txtErrPath.SetFocus: GoTo Warning
        IsDataOk = True
    
    Exit Function
Warning:
        MsgBox "��� " & strErrMsg & " ���ݦ���", vbOKOnly + vbInformation, "�T��"
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set rsTmp = Nothing
    If rsCD013.State = adStateOpen Then rsCD013.Close
    Set rsCD013 = Nothing
    file(0).Close
    Set fso = Nothing
End Sub

Private Sub gdaSystemDate_LostFocus()
   Dim intDD As Integer
   If Len(gdaSystemDate.Text) <> 0 Then
   intDD = CInt(Format(gdaSystemDate.Text, "DD"))
   txtPayDate.Text = Format(intDD, "00")
   End If
End Sub

''2005/01/26�N��z�n�����  �̷ӳ�ڱƧǤ���  �N�P�渹���t���X��
Private Sub LogSameNOPMAmount(ByVal strMNO As String, ByVal CardName As String, ByVal set34sql As String)

    Dim checkSameNOSQL As String
    Dim snoType As String
    Dim rscheckSameNOSQL As New ADODB.Recordset
    
    If intPara24 = 0 Then
         snoType = "BillNO"
    Else
         snoType = "MediaBillNo"
    End If
    
checkSameNOSQL = "SELECT Count(A." & snoType & ")   " & _
                    "From " & _
                    TableOwnerName & "SO033 A ," & _
                    TableOwnerName & "SO001 B ," & _
                    TableOwnerName & "SO002A C,  " & _
                    TableOwnerName & "SO002 D,  " & _
                    TableOwnerName & "SO106 E  " & _
                    "Where " & strBankId & " AND  " & _
                    "C.CARDNAME  ='" & CardName & "' AND " & _
                    "A.CustID  = C.CUSTID AND " & _
                    "A.CustID  = D.CUSTID AND " & _
                    "A.CustID  = E.CUSTID AND C.AccountNO  = E.AccountID  AND  " & _
                    "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                    "A.UCCode Is Not Null AND " & _
                    "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo  AND E.StopFlag = 0  " & _
                    IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34sql & "  AND A." & snoType & " = '" & strMNO & "'  AND A.ShouldAmt< 0   "
    
If GetRS(rscheckSameNOSQL, checkSameNOSQL, cn, adUseClient, adOpenKeyset) = True Then
      If rscheckSameNOSQL(0) > 0 Then
                '' �i��  log �O��
                file(0).WriteLine " �t�����B : �渹�@" & strMNO
      End If
End If

rscheckSameNOSQL.Close
Set rscheckSameNOSQL = Nothing

Exit Sub

ChkErr:
    Call ErrSub(Me.Name, "LogSameNOPMAmount")
End Sub

Private Function GetSumAmout(ByVal strMNO As String, ByVal CardName As String, ByVal set34sql As String)
    Dim SingalSumAmount As Long: SingalSumAmount = 0
    Dim checkSameNOSQL As String
    Dim snoType As String
    Dim rscheckSameNOSQL As New ADODB.Recordset
    
        If intPara24 = 0 Then
             snoType = "BillNO"
        Else
             snoType = "MediaBillNo"
        End If
         '' �N C.StopFlag = 0  �o�Ӯ����A�קK�j������
        '#4040 �h��SO001 ���M��ܫȤ����O�|�X�� By Kin 2008/08/13
        checkSameNOSQL = "SELECT SUM(A.ShouldAmt)  ShouldAmt  " & _
                            "From " & _
                            TableOwnerName & "SO033 A," & _
                            TableOwnerName & "SO002 D," & _
                            TableOwnerName & "SO001 B " & _
                            "Where " & _
                            "A.CancelFlag = 0 And " & _
                            "A.CustID  = D.CUSTID AND " & _
                            "A.SERVICETYPE = D.SERVICETYPE AND " & _
                            "A.UCCode Is Not Null And " & _
                            "A.CustId=B.CustId " & _
                            IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34sql & "  AND A." & snoType & " = '" & strMNO & "'   " & _
                            " Group by A." & snoType
        
        If GetRS(rscheckSameNOSQL, checkSameNOSQL, cn, adUseClient, adOpenKeyset) = True Then
            SingalSumAmount = rscheckSameNOSQL(0)
        Else
            file(0).WriteLine " ���D���B : �渹�@" & strMNO
        End If
        
        rscheckSameNOSQL.Close
        Set rscheckSameNOSQL = Nothing
        GetSumAmout = SingalSumAmount
    Exit Function

ChkErr:
    Call ErrSub(Me.Name, "GetSumAmout")
End Function

Private Function GetStopYMDateString(ByVal strMNO As String, ByVal CardName As String, ByVal set34sql As String) As String
    Dim rsStopYMDateString As New ADODB.Recordset
    Dim rsStopYMDateSQL As String
    Dim strCardNameSQL As String
        If CardName <> "" Then strCardNameSQL = " And C.CARDNAME  ='" & CardName & "'"
        
        If intPara24 = 0 Then
           rsStopYMDateSQL = "SELECT E.StopYM   " & _
                    "From " & _
                    TableOwnerName & "SO033 A ," & _
                    TableOwnerName & "SO001 B ," & _
                    TableOwnerName & "SO002A C,  " & _
                    TableOwnerName & "SO002 D,  " & _
                    TableOwnerName & "SO106 E  " & _
                    "Where " & strBankId & strCardNameSQL & " AND " & _
                    "A.CustID  = C.CUSTID AND " & _
                    "A.CustID  = D.CUSTID AND " & _
                    "A.CustID  = E.CUSTID AND C.AccountNO  = E.AccountID  AND  " & _
                    "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                    "A.UCCode Is Not Null  AND  " & _
                    "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo AND E.StopFlag = 0 " & _
                    IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34sql & "  AND A.BillNO = '" & strMNO & "'   " & _
                    " Group by A.BillNO , E.AccountID, E.StopYM  ORDER BY  A.BillNO,SUBSTR(LPAD(E.STOPYM,6,'0'),3,4) || SUBSTR(LPAD(E.STOPYM,6,'0'),1,2) DESC  "
    
        Else
          rsStopYMDateSQL = "SELECT  E.StopYM    " & _
                    "From " & _
                    TableOwnerName & "SO033 A ," & _
                    TableOwnerName & "SO001 B ," & _
                    TableOwnerName & "SO002A C,  " & _
                    TableOwnerName & "SO002 D,  " & _
                    TableOwnerName & "SO106 E  " & _
                    "Where " & strBankId & strCardNameSQL & " AND " & _
                    "A.CustID  = C.CUSTID AND " & _
                    "A.CustID  = D.CUSTID AND " & _
                    "A.CustID  = E.CUSTID AND C.AccountNO  = E.AccountID  AND  " & _
                    "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                    "A.UCCode Is Not Null AND  " & _
                    "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo  AND E.StopFlag = 0  " & _
                    IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34sql & "  AND A.MediaBillNo= '" & strMNO & "'   " & _
                    " Group by  E.AccountID , E.StopYM ,  A.MediaBillNo  ORDER BY  A.MediaBillNo ,SUBSTR(LPAD(E.STOPYM,6,'0'),3,4) || SUBSTR(LPAD(E.STOPYM,6,'0'),1,2) DESC  "
    
                    
        End If
        
        If GetRS(rsStopYMDateString, rsStopYMDateSQL, cn, adUseClient, adOpenDynamic) = True Then
            If IsNull(rsStopYMDateString(0)) Then
                GetStopYMDateString = "NULL"
            Else
                GetStopYMDateString = rsStopYMDateString(0)
            End If
        Else
            GetStopYMDateString = "NULL"
        End If
    
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "GetStopYMDateString")
End Function
Function subCitiBankBatch() As Boolean
    On Error GoTo ChkErr
    Dim intTxtFileIndex As Integer
    Dim lngSubCount As Long
    Dim lngTotalAmount As Long
    Dim strFileName As String
        lngSubCount = 0
        lngTotalAmount = 0
        intTxtFileIndex = 1
        Call DefineRs3    ''�إ߰O���� recordset ���c
        If Not BeginTran("", txtSpcNO, txtMerchantName, intTxtFileIndex, lngSubCount, lngTotalAmount, True) Then
            'Exit Function
        End If
        If lngSubCount > 0 Then
            strFileName = txtDataPath & "\" & txtSpcNO & Format(Date, "yymmdd") & "." & txtBatch
            objStorePara.uProcText = strFileName
            Set file(intTxtFileIndex) = fso.CreateTextFile(strFileName, True)
            Dim lngSetCount(4) As Long, lngSetTotalAmt(4) As Long
            strHeader = "FHD"
            strHeader = strHeader & Right(Format(Date, "YYYYMMDD"), 6)
            strHeader = strHeader & GetString(txtBatch, 2, GIRIGHT, True)
            strHeader = strHeader & GetString(lngSubCount, 8, GIRIGHT, True)
            strHeader = strHeader & GetString(lngTotalAmount, 13, GIRIGHT, True) & "00"
            '*****************************************************************************
            '#3305 �NHead�O�d�����ץѭ�Ӫ�116�אּ183 By Kin 2007/07/26
            strHeader = strHeader & Space(183)
            '*****************************************************************************
            Call subWriteLine(intTxtFileIndex)   ''''' �N���o����Ƽg�J��r��
        End If
        subCitiBankBatch = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subCitiBankBatch")
End Function
Private Sub DefineRs3()
    On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        
        ''"""""""""""""""""""""""""""""""""
        .Fields.Append ("F01"), adBSTR, 3, adFldIsNullable
        .Fields.Append ("F02"), adBSTR, 3, adFldIsNullable
        .Fields.Append ("F03"), adBSTR, 12, adFldIsNullable
        .Fields.Append ("F04"), adBSTR, 19, adFldIsNullable
        .Fields.Append ("F05"), adBSTR, 2, adFldIsNullable
        
        .Fields.Append ("F06"), adBSTR, 2, adFldIsNullable
        .Fields.Append ("F07"), adBSTR, 12, adFldIsNullable
        .Fields.Append ("F08"), adBSTR, 13, adFldIsNullable
        .Fields.Append ("F09"), adBSTR, 16, adFldIsNullable
        .Fields.Append ("F10"), adBSTR, 3, adFldIsNullable
        
        .Fields.Append ("F11"), adBSTR, 6, adFldIsNullable
        .Fields.Append ("F12"), adBSTR, 12, adFldIsNullable
        .Fields.Append ("F13"), adBSTR, 2, adFldIsNullable
        .Fields.Append ("F14"), adBSTR, 4, adFldIsNullable
        .Fields.Append ("F15"), adBSTR, 8, adFldIsNullable
        
        .Fields.Append ("F16"), adBSTR, 8, adFldIsNullable
        .Fields.Append ("F17"), adBSTR, 15, adFldIsNullable
        .Fields.Append ("F18"), adBSTR, 6, adFldIsNullable
        .Fields.Append ("F19"), adBSTR, 2, adFldIsNullable
        .Fields.Append ("F20"), adBSTR, 2, adFldIsNullable
        '*********************************************************************
        '#3305 �s�W�H1�ӫO�d������67 By Kin 2007/07/26
        .Fields.Append ("F21"), adBSTR, 67, adFldIsNullable
        '*********************************************************************
        .Open
    End With
    
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefineRs3")
End Sub
Private Function BeginTran3(ByRef lngSubCount As Long, ByRef lngTotal As Long, _
    ByVal strCardExpDate As String, theSingalAmount As Long, ByVal strSpcNo As String) As Boolean
    
    On Error GoTo ChkErr
    Dim strBillNO As String
    If intPara24 = 0 Then
         strBillNO = "BillNO"
    Else
         strBillNO = "MediaBillNO"
    End If
    With rsDefTmp
        .AddNew
        'A Detail Record�T�w��"FDT"
        .Fields("F01").Value = "FDT"
        'N �����w�N���AMerchantID(�̦U�����ڸ�Ƥ�SO106.CardCode������CD037.CodeNo��CD037.SPCNO)�e�T�X
        .Fields("F02").Value = Right("000" & Trim(GetString(strSpcNo, 3)), 3)
        '�ť�
        .Fields("F03").Value = Space(12)
        '�q��s��(�C��渹�Φ��O�渹�̾ڦ��O�ѼƬO�_�Ұ�'�C��h�b��B�z')
        .Fields("F04").Value = GetString(rsTmp(strBillNO), 19)
        'N �������ơA���@�����h��ť�(�ثe��ť�)
        .Fields("F05").Value = Space(2)
        '�ť�
        .Fields("F06").Value = Space(2)
        '�Ȱ�Gauth: ���v���\ authfail: ���v���� sale: �ʪ����\ salefail: �ʪ����� cap: �дڦ��\ capfail: �дڥ��� settle: �дڵ��b
        .Fields("F07").Value = Space(12)
        'N �t�G��p�Ʀ�(SO033.SHOULDAMT)
        .Fields("F08").Value = GetString(theSingalAmount, 11, GIRIGHT, True) & "00"
        lngTotal = lngTotal + theSingalAmount  '�`���B=���B�֥[
        .Fields("F09").Value = GetString(rsTmp("AccountNo"), 16, GIRIGHT)
        '**************************************************************
        '#3305 ��10�����쥻�O��000�A�{�b�אּ3�Ӫť� By Kin 2007/07/26
        .Fields("F10").Value = Space(3)
        '**************************************************************
        .Fields("F11").Value = Format(strCardExpDate, "YYYYMM")
        .Fields("F12").Value = Space(12)
        .Fields("F13").Value = Space(2)
        .Fields("F14").Value = Space(4)
        .Fields("F15").Value = Space(8)
        .Fields("F16").Value = GetString(Trim(txtMerchantName), 8, GIRIGHT, True)
        .Fields("F17").Value = GetString(strSpcNo, 15)
        .Fields("F18").Value = Space(6)
        .Fields("F19").Value = Space(2)
        .Fields("F20").Value = Space(2)
        '*****************************************************
        '#3305�W�[�@�ӫO�d���A����67 By Kin 2007/07/26
        .Fields("F21").Value = Space(67)
        '*****************************************************
        .Update
    End With
    BeginTran3 = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "BeginTran3"
End Function


