VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCitiBank 
   BorderStyle     =   1  '��u�T�w
   ClientHeight    =   5115
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
   ScaleHeight     =   5115
   ScaleWidth      =   9240
   StartUpPosition =   1  '���ݵ�������
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   7920
      Top             =   915
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
      TabIndex        =   13
      Top             =   345
      Width           =   8895
      Begin VB.CheckBox chkDuteDate 
         Caption         =   "�H�Υd�L����Ƥ@�ֲ���"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2040
         TabIndex        =   5
         Top             =   1995
         Width           =   2535
      End
      Begin VB.TextBox txtSpcNO 
         Height          =   300
         Left            =   2025
         MaxLength       =   9
         TabIndex        =   24
         Top             =   435
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.TextBox txtMerchantName 
         Height          =   300
         Left            =   6060
         MaxLength       =   9
         TabIndex        =   12
         Text            =   "TBGB"
         Top             =   390
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.TextBox txtStatement 
         Height          =   345
         Left            =   2025
         TabIndex        =   4
         Top             =   1530
         Width           =   6465
      End
      Begin VB.TextBox txtDiscountRate 
         Height          =   345
         Left            =   2025
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "0"
         Top             =   1140
         Width           =   900
      End
      Begin VB.TextBox txtPayDate 
         Height          =   345
         Left            =   6075
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "00"
         Top             =   1140
         Width           =   330
      End
      Begin VB.ComboBox cboSet 
         Height          =   315
         ItemData        =   "frmCitiBank.frx":0442
         Left            =   2025
         List            =   "frmCitiBank.frx":0452
         TabIndex        =   0
         Top             =   780
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
         TabIndex        =   14
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
            TabIndex        =   6
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
            Left            =   7335
            TabIndex        =   9
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
            TabIndex        =   7
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
            TabIndex        =   8
            Top             =   765
            Width           =   4095
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "����ɦ�m�]���| )"
            Height          =   195
            Left            =   1290
            TabIndex        =   16
            Top             =   375
            Width           =   1665
         End
         Begin VB.Label lblErrpath 
            AutoSize        =   -1  'True
            Caption         =   "���D�Ѧ��ɦ�m�]���| + �W�١^"
            Height          =   195
            Left            =   330
            TabIndex        =   15
            Top             =   810
            Width           =   2640
         End
      End
      Begin Gi_Date.GiDate gdaSystemDate 
         Height          =   345
         Left            =   6060
         TabIndex        =   2
         Top             =   720
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
      Begin VB.Label Label7 
         Alignment       =   1  '�a�k���
         Caption         =   "�H�Υd���b�S���N�X"
         Height          =   285
         Left            =   45
         TabIndex        =   25
         Top             =   510
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label5 
         Alignment       =   1  '�a�k���
         Caption         =   "�t�Τ��"
         Height          =   285
         Left            =   4215
         TabIndex        =   23
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label Label9 
         Caption         =   "(DD)"
         Height          =   240
         Left            =   6465
         TabIndex        =   22
         Top             =   1170
         Width           =   1080
      End
      Begin VB.Label Label8 
         Caption         =   "�H�Υd���b�S���W��"
         Height          =   285
         Left            =   4215
         TabIndex        =   21
         Top             =   435
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "�Ȥ�b�椺�e"
         Height          =   210
         Left            =   750
         TabIndex        =   20
         Top             =   1620
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  '�a�k���
         Caption         =   "�馩�v"
         Height          =   210
         Left            =   1215
         TabIndex        =   19
         Top             =   1245
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "�дڤ��"
         Height          =   210
         Left            =   5190
         TabIndex        =   18
         Top             =   1185
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "���b����"
         Height          =   285
         Left            =   1140
         TabIndex        =   17
         Top             =   870
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�}�l"
      Height          =   375
      Left            =   315
      TabIndex        =   10
      Top             =   4470
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   7665
      TabIndex        =   11
      Top             =   4455
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

Dim rsCD037SpcNO  As New ADODB.Recordset  '' ���o SPCNO ���ߤ@�Ȫ��O���A�H�P��SPCNO���d�O�@�����X
Dim rsCD037 As New ADODB.Recordset
Dim s As String
Dim intTxtFileIndex  As Integer
Dim lngTime As Long


lngTime = Timer
 If Not IsDataOk Then Exit Sub
 '' Call DefineRs   ''�إ߰O���� recordset ���c
 
  s = txtDataPath.Text
  If Not CreateObject("Scripting.FileSystemObject").FolderExists(s) Then MsgBox "���| " & s & " ���s�b!", vbExclamation: Exit Sub

  If OpenFile(txtErrPath, 0, True) = False Then Exit Sub
  
  ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
  ' ���o SPCNO ���ߤ@�Ȫ��O���A�H�P��SPCNO���d�O�@�����X
  ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
  rsCD037SpcNO.CursorLocation = adUseClient
  rsCD037SpcNO.Open _
      "SELECT DISTINCT SPCNO FROM " & TableOwnerName & "CD037  WHERE SpcNO IS NOT NULL  ", _
      cn, adOpenKeyset, adLockReadOnly
  If rsCD037SpcNO.EOF Or rsCD037SpcNO.BOF Then
        MsgBox "�H�Υd���������b�S���N�X���ݳ]�� "
        rsCD037SpcNO.Close
        Exit Sub
  End If

 ' ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
 
  rsCD037.CursorLocation = adUseClient
  
  rsCD037.Open _
          "SELECT * FROM " & TableOwnerName & "CD037 " & _
          "WHERE " & _
          "Spcno IS NOT NULL AND MerchantName IS NOT NULL " & _
          "ORDER BY CodeNO", _
          cn, adOpenKeyset, adLockReadOnly
          
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
                     ''MsgBox " �H�Υd�O�� " & rsCD037("Description") & "�@�����X�@!!", vbOKOnly + vbCritical, "�T��"
                End If
                rsCD037.MoveNext
                blnNewRec = True
     Loop
     
     '' �N�����ұo����ƶ�J��r��
     
     If lngSubCount > 0 Then
           rsCD037.MoveFirst
 ''       If OpenFile(txtDataPath, intTxtFileIndex, True, cboSet.ListIndex, rsCD037("MerchantName"), rsCD037("SPCNO")) = False Then

            If OpenFile(txtDataPath, intTxtFileIndex, True, cboSet.ListIndex, rsCD037("MerchantName"), rsCD037("CardID")) = False Then
                MsgBox "�L�k�إߵn���S���N�X�� " & rsCD037("SPCNO") & " ����r�� !!", vbOKOnly + vbCritical, "�T�� "
                Exit Sub
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

Call CloseFS(intTxtFileIndex)
Call ScrToRcd
objStorePara.uUpdate = True

rsCD037.Close
Set rsCD037 = Nothing
rsCD037SpcNO.Close
Set rsCD037SpcNO = Nothing
If lngErrCount > 0 Then Call Shell("NotePad " & txtErrPath.Text, vbNormalFocus)
Call msgResult(lngCount, lngErrCount, lngTime)      '��ܰ��浲�G
Unload Me

Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub Form_Load()

    On Error Resume Next
    blnSameFile = False
    lngErrCount = 0
    lngCount = 0
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
            
              '' If Not .AtEndOfStream Then txtMerchantName.Text = .ReadLine & ""   '�H�Υd���b�S���W��
              '' If Not .AtEndOfStream Then txtPayDate.Text = .ReadLine & ""   '�дڤ��
               If Not .AtEndOfStream Then txtDiscountRate.Text = .ReadLine & ""   '�馩�v
               If Not .AtEndOfStream Then txtStatement.Text = .ReadLine & ""  '�Ȥ�b�椺�e
'              If Not .AtEndOfStream Then txtInVoice.Text = .ReadLine & ""   '�Τ@�s��

               If Not .AtEndOfStream Then txtDataPath = .ReadLine & ""   '�����
               If Not .AtEndOfStream Then txtErrPath = .ReadLine & ""   '���D��
                       
                  
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
  
  
  ''2005/01/10  �[�J�P�_��  ���H�Υd���Ħ~��̤j��
  Dim strAccountID As String
  Dim lngCustID As Long
  Dim CountSameCustIDBillNO As Integer
  
  Dim theSingalAmount As Long  '' 2005/02/02 �Ψ��x�s�浧SO033��ơA
  
  
  Dim lngerror As Long:    lngerror = 0
  
  strAccountID = ""
  lngCustID = 0
  CountSameCustIDBillNO = 0
  
  
     
  BeginTran = False
    
'' 2005/02/02 ���ޥ��t���F  �@�߳��X��
    
'   If cboSet.ListIndex = 2 Then
'       set34 = "  AND A.ShouldAmt > 0 "
'   Else
'        set34 = "  AND A.ShouldAmt < 0 "
'   End If
   set34 = "  "
   
   ''' 2003/07/23 �s�W�Ȥ�s��(CustID)�H�δC��渹(MediaBillNo)�W��
   ''' 2003/08/14 �令SO002A ����ƨӷ��A�åB�NGROUP BY ���r�y�ζ}�A�����ϥδC��渹�P���ϥΪ�
   '''�ϥΥHMediaBillNO ���D�A�S���H Billno���D �A�������Ȥ�s����T �h�H  ���s���o��SO002A ���ǡACustStatusCode  �b SO002A �S�� �]���n�N SO002 ��i��

'       strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
'                "A.BillNO, sum(A.ShouldAmt)  ShouldAmt,  " & _
'                "C.AccountNO,C.CardExpDate,B.CUSTNAME ,A.MediaBillNo,A.CUSTID   " & _
'                "From " & _
'                TableOwnerName & "SO033 A ," & _
'                TableOwnerName & "SO001 B ," & _
'                TableOwnerName & "SO002A C  " & _
'                "Where " & strBankId & " AND  " & _
'                "C.CARDNAME  ='" & strCardName & "' AND " & _
'                "A.CustID  = C.CUSTID AND " & _
'                "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
'                "A.UCCode Is Not Null And A.ShouldAmt > 0 " & _
'                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
'                " Group by A.BillNO , C.AccountNO , C.CardExpDate , B.CUSTNAME, A.MediaBillNo,A.CustID  "

''  ���q�� 2004/08/12 �O�d�A�䤤�� C.CardExpDate  �令�� SO106 �� StopYM
'' *****************************************************************************************************
'  If intPara24 = 0 Then
'       strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
'                "A.BillNO, sum(A.ShouldAmt)  ShouldAmt,  " & _
'                "C.AccountNO,C.CardExpDate   " & _
'                "From " & _
'                TableOwnerName & "SO033 A ," & _
'                TableOwnerName & "SO001 B ," & _
'                TableOwnerName & "SO002A C,  " & _
'                TableOwnerName & "SO002 D  " & _
'                "Where " & strBankId & " AND  " & _
'                "C.CARDNAME  ='" & strCardName & "' AND " & _
'                "A.CustID  = C.CUSTID AND " & _
'                "A.CustID  = D.CUSTID AND " & _
'                "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
'                "A.UCCode Is Not Null And A.ShouldAmt > 0  AND " & _
'                "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo  " & _
'                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
'                " Group by A.BillNO , C.AccountNO , C.CardExpDate  "
'    Else
'      strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
'                "sum(A.ShouldAmt)  ShouldAmt,  " & _
'                "C.AccountNO,C.CardExpDate ,A.MediaBillNo   " & _
'                "From " & _
'                TableOwnerName & "SO033 A ," & _
'                TableOwnerName & "SO001 B ," & _
'                TableOwnerName & "SO002A C,  " & _
'                TableOwnerName & "SO002 D  " & _
'                "Where " & strBankId & " AND  " & _
'                "C.CARDNAME  ='" & strCardName & "' AND " & _
'                "A.CustID  = C.CUSTID AND " & _
'                "A.CustID  = D.CUSTID AND " & _
'                "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
'                "A.UCCode Is Not Null And A.ShouldAmt > 0  AND " & _
'                "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo  " & _
'                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
'                " Group by  C.AccountNO , C.CardExpDate ,  A.MediaBillNo "
'    End If

'' *****************************************************************************************************
    ''  2004/09/24 �[�JSO106.StopFlag = 0 ������� ����X�{�ⵧ
    ''  2004/09/24 �٬O�N��qmark���n�F�A������SO106�N�n�F
    ''  �H�U����qSQL �N�䤤�� A.ShouldAmt > 0  �����󮳱�
    
    If intPara24 = 0 Then
       strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                "A.BillNO, sum(A.ShouldAmt)  ShouldAmt,  " & _
                "E.AccountID  AccountNO,E.StopYM CardExpDate  " & _
                "From " & _
                TableOwnerName & "SO033 A ," & _
                TableOwnerName & "SO001 B ," & _
                TableOwnerName & "SO002A C,  " & _
                TableOwnerName & "SO002 D,  " & _
                TableOwnerName & "SO106 E  " & _
                "Where " & strBankId & " AND  " & _
                "C.CARDNAME  ='" & strCardName & "' AND " & _
                "A.CustID  = C.CUSTID AND " & _
                "A.CustID  = D.CUSTID AND " & _
                "A.CustID  = E.CUSTID AND C.AccountNO  = E.AccountID  AND  " & _
                "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                "A.UCCode Is Not Null  AND  " & _
                "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo AND E.StopFlag = 0 " & _
                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
                " Group by A.BillNO , E.AccountID, E.StopYM  ORDER BY  A.BillNO,E.StopYM DESC  "
        '        " Group by A.BillNO , E.AccountID, E.StopYM  ORDER BY  A.BillNO,SUBSTR(LPAD(E.STOPYM,6,'0'),3,4) || SUBSTR(LPAD(E.STOPYM,6,'0'),1,2) DESC  "
          
               ' " Group by A.BillNO , C.AccountNO , E.StopYM  "  ''�o�@�q20050110 �ձ������F��SO106 �����
    Else
      strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                "sum(A.ShouldAmt)  ShouldAmt,  " & _
                "E.AccountID  AccountNO,E.StopYM CardExpDate  ,A.MediaBillNo   " & _
                "From " & _
                TableOwnerName & "SO033 A ," & _
                TableOwnerName & "SO001 B ," & _
                TableOwnerName & "SO002A C,  " & _
                TableOwnerName & "SO002 D,  " & _
                TableOwnerName & "SO106 E  " & _
                "Where " & strBankId & " AND  " & _
                "C.CARDNAME  ='" & strCardName & "' AND " & _
                "A.CustID  = C.CUSTID AND " & _
                "A.CustID  = D.CUSTID AND " & _
                "A.CustID  = E.CUSTID AND C.AccountNO  = E.AccountID  AND  " & _
                "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                "A.UCCode Is Not Null AND  " & _
                "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo  AND E.StopFlag = 0  " & _
                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
                " Group by  E.AccountID , E.StopYM ,  A.MediaBillNo  ORDER BY  A.MediaBillNo ,E.StopYM DESC  "
          '     " Group by  E.AccountID , E.StopYM ,  A.MediaBillNo  ORDER BY  A.MediaBillNo ,SUBSTR(LPAD(E.STOPYM,6,'0'),3,4) || SUBSTR(LPAD(E.STOPYM,6,'0'),1,2) DESC  "
           '
               ' " Group by  C.AccountNO , E.StopYM ,  A.MediaBillNo  "  ''�o�@�q20050110 �ձ������F��SO106 �����
                
    End If
    
    
    
 ''  2003/08/14 �令 �A�N mediabillno   ��ƨ��X �o�سo�@���h���@�� BillNO �M���J��mediabillno��null �Ȫ��ɭԡA��J�C��渹
    If intPara24 = 1 Then
    
                        ''�o�@�q�������e�z���D SQL (strsql)  ���� group  �A�[�W Billno
                        
                        strMediaBillNONull = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                                                           "C.AccountNO,A.MediaBillNo,A.BillNO   " & _
                                                            "From " & _
                                                            TableOwnerName & "SO033 A ," & _
                                                            TableOwnerName & "SO001 B ," & _
                                                            TableOwnerName & "SO002A C,  " & _
                                                            TableOwnerName & "SO002 D  " & _
                                                            "Where " & strBankId & " AND  " & _
                                                            "C.CARDNAME  ='" & strCardName & "' AND " & _
                                                            "A.CustID  = C.CUSTID AND " & _
                                                            "A.CustID  = D.CUSTID AND " & _
                                                            "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                                                            "A.UCCode Is Not Null  AND " & _
                                                            "A.SERVICETYPE=D.SERVICETYPE  AND A.AccountNo = C.AccountNo  " & _
                                                            IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & "  AND A.MEDIABILLNO IS NULL"
                                                            
                                                            
                            strChooseMultiAcc = " A.CancelFlag = 0 AND " & _
                                                              "A.UCCode Is Not Null  " & _
                                                              IIf(Len(strChooseMultiAcc) = 0, "", " And ") & strChooseMultiAcc
                
                        If GetRS(rsMedioBillnoNull, strMediaBillNONull, cn) Then
                                     Do While Not rsMedioBillnoNull.EOF
                                          If IsNull(rsMedioBillnoNull("MediaBillNo")) Then
                                                  strSequenceNumber = GetInvoiceNo2("SO033_MediaBillNo", cn)
                                                  cn.Execute ( _
                                                                            "UPDATE " & TableOwnerName & "SO033 A SET " & _
                                                                            "A.MediaBillNo  ='" & strSequenceNumber & "'  WHERE   " & _
                                                                            "A.BillNo = '" & rsMedioBillnoNull("BillNO") & "'   AND  " & strChooseMultiAcc & " AND " & _
                                                                            "A.AccountNO ='" & rsMedioBillnoNull("AccountNO") & "' ")
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
    Else
        blnNodata = False
        Dim strSN As String  '' ���w�ǦC��
        Dim lngtotal As Long '' �֭p���Y�� Total Amount ����`�X����B
        Dim lngDoubleMediaNO As Long '' �o�ӭȰO�����ХX�檺��l���

        lngDoubleMediaNO = 0
        If pNewRec = False Then
             lngtotal = 0
             lngSubCount = 0
        Else
            lngtotal = pTotalAmount
            lngSubCount = pSubCount
        End If
                                  
        While Not rsTmp.EOF
        
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
                strSopYMMaxString = GetStopYMDateString(strAccountID, strCardName, set34)
                lngerror = 45
           End If
            '' ************************************************************************
      '     ''2005/01/26�N��z�n�����  �̷ӳ�ڱƧǤ���  �N�P�渹���t���X��
           '''Call LogSameNOPMAmount(rsTmp(strBillNOField), strCardName, set34)
           ''2005/02/02 �N��z�n�����  �̷ӳ�ڱƧǤ���  �N�P�渹���t���X��
           
           lngerror = 5
             Call GetRS(rsCustIDErr, "SELECT  SO001.CUSTNAME , SO002A.CUSTID FROM " & TableOwnerName & "SO002A," & TableOwnerName & "SO001  " & _
                                  "WHERE " & _
                                  "SO002A.AccountNO ='" & rsTmp("AccountNO") & "'  AND  SO002A.CUSTID = SO001.CUSTID ", cn)
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
             strCardExpDate = Right(strSopYMMaxString, 4) & "/" & Replace(strSopYMMaxString, Right(strSopYMMaxString, 4), "")
             lngerror = 10
           
           dteCED = CDate(strCardExpDate)
           If dteCED < CDate(Format(gdaSystemDate.Text, "YYYY/MM")) Then
               lngErrCount = lngErrCount + 1
               strErrMessage = "�H�Υd�L�� : �渹�@" & rsTmp.Fields(strBillNOField) & " �Ȥ�m�W�@" & rsCustIDErr("custname") & " �H�Υd�� " & rsTmp.Fields("AccountNO") & " �����@" & rsTmp.Fields("CardExpDate") & " �Ȥ�s��  " & rsCustIDErr("CustID")
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
        ''   Else
               lngCount = lngCount + 1
'               lngtotal = lngtotal + rsTmp.Fields("ShouldAmt")
               lngSubCount = lngSubCount + 1
               With rsDefTmp
                    .AddNew
                    .Fields("SequenceNo") = Format(lngSubCount, "000000")
                    '' SequenceNo(1-6)
                    '' .Fields("MerchantNo") = Format(txtSpcNO.Text, "000000000")
                    .Fields("MerchantNo") = Format(strSpcNo, "000000000")
                    ''MerchantNo(7-15)
                    .Fields("DiscountRate") = Format(txtDiscountRate.Text, "00000")
                    ''DiscountRate(16-20)
                    .Fields("CardNo") = IIf(Len(rsTmp("AccountNO")) > 0, Format(rsTmp("AccountNO"), "0000000000000000"), "0000000000000000")
                    ''CardNo(21-36)
                  ''  .Fields("ExpiryDate") = IIf(IsNull(rsTmp("CardExpDate")), "000000", Format(rsTmp("CardExpDate"), "000000"))
                  '' 20050205 MARK�H�W�����@��A�令�H�U��
                  strCardExpDate = Replace(strCardExpDate, "/", "")
                  If Len(strCardExpDate) = 5 Then
                      strCardExpDate = Left(strCardExpDate, 4) & "0" & Right(strCardExpDate, 1)
                  End If
                  ''.Fields("ExpiryDate") = Replace(strCardExpDate, "/", "")
                  .Fields("ExpiryDate") = Right(strCardExpDate, 2) & Left(strCardExpDate, 4)
                    ''ExpiryDate(37-42)
                    .Fields("SequenceNo2") = Format(lngSubCount, "000000000000000")
                    ''SequenceNo2(43-57)
                    .Fields("Filler1") = Space(8)
                    '' Filler(58-65)
                    ''2005/02/02 �H�U���B���ȧ令����SO033�����B �קK�{�����o���Ъ��B  !!
                    ''.Fields("Amount") = Format(rsTmp("ShouldAmt"), "0000000")
                    .Fields("Amount") = Format(theSingalAmount, "0000000")
                    lngtotal = lngtotal + theSingalAmount
                    '' Amount(66-72)
                     Dim strSD As String
                     Dim intlength As Integer
                     
                     '""�ഫ�������T���ת����Φr��'''''''''''''''''''''''''
                     strSD = StrConv(txtStatement.Text, vbFromUnicode)
                     intlength = LenB(strSD)
                     'strSD = txtStatement.Text & IIf(intlength < 38, Space(38 - intlength), "")
                     strSD = txtStatement.Text & Space(38)
                   '  MsgBox LenB(StrConv(strSD, vbFromUnicode))
                     
                 ''  strSD = StrConv(LeftB(StrConv(strSD, vbFromUnicode), 38), vbUnicode)
                     ''""""""""""""""""""""""""""""""""""""""""""""""""""""
                   .Fields("StatementDescription") = StrConv(LeftB(StrConv(StrConv(strSD, vbWide), vbFromUnicode), 38), vbUnicode)
                    
                    '.Fields("StatementDescription") = txtStatement.Text & Space(38 - Len(txtStatement.Text))
                    ''StatementDescription(73-110)
                    .Fields("Filler2") = Space(2)
                    
                   '' 2003/07/23 �H�U�i��渹��Ȫ��ާ@�A�P�_�O�_�H�C��渹���\��J��
                   '' �p�G�O�h�H�C��渹�J��A�_�h�ഫ���³渹����A�J
                   '' 2003/08/14 �H�U�����e�A���o�渹���e�A�p�G�ҥδC�骺�\��A�N������MediaBillNO
                   
                   If intPara24 = 0 Then
                             Dim strBillNOOld As String
                             strBillNOOld = Trim(CStr(Val(Left(rsTmp("BillNO"), 6)) - 191100)) & _
                                            Mid(rsTmp("BillNO"), 7, 1) & Format(Right(rsTmp("BillNO"), 6), "000000")
                            .Fields("MerchantDescription") = strBillNOOld & Space(20 - Len(strBillNOOld))
                   Else
'                          If Not IsNull(rsTmp("MediaBillNo")) Then
'                          lngDoubleMediaNO = lngDoubleMediaNO + 1
'                             .Fields("MerchantDescription") = GetString(rsTmp("MediaBillNo"), 20, giLeft, False)
                             
                       ''     2003/08/06   ���������{���X���A�ݭn�ϥΤF�A�]���@�w�|�C��渹
                       ''      file(0).WriteLine "�渹   [" & rsTmp("BillNo") & "]  ���ХX��F" & _
                                                       "�C��渹 [" & rsTmp("MediaBillNo") & "]�F" & _
                                                       "�Ȥ�s�� [" & rsTmp("CustID") & "]  "
'                          Else
'                              strSequenceNumber = GetInvoiceNo2("SO033_MediaBillNo", cn)
                          ''    .Fields("MerchantDescription") = GetString(strSequenceNumber, 20, giLeft, False)
                          '' ������MediaBillNO
                           .Fields("MerchantDescription") = GetString(rsTmp("MediaBillNo"), 20, giLeft, False)
                    ''         cn.Execute ("UPDATE SO033 A SET MediaBillNo  ='" & strSequenceNumber & "'  WHERE   BillNo = '" & rsTmp("BillNO") & "'   AND  " & strChooseMultiAcc & " AND CustID = " & rsTmp("CUSTID") & " AND AccountNO ='" & rsTmp("AccountNO") & "'")
                   ''    End If
                   End If
                   
                   '' �o�@�q���ѬOemc �� �K�ӰѦҥ�
'                                     If intPara24 = 0 Then
'                                             .Fields("F97112") = GetString(rsTmp("BillNO"), 16, giLeft, False)
'                                    Else
'                                         If Not IsNull(rsDoubleMediaNO("MediaBillNo")) Then
'                                            lngDoubleMediaNO = lngDoubleMediaNO + 1
'                                            .Fields("F97112") = GetString(rsDoubleMediaNO("MediaBillNo"), 16, giLeft, False)
'                                             file(1).WriteLine "�渹   [" & rsTmp("BillNo") & _
'                                             "]  ���ХX��A�C��渹 [" & rsDoubleMediaNO("MediaBillNo") & _
'                                             "] �Ȥ�s�� [" & rsTmp("CustID") & "]  "
'                                        Else
'                                             strSequenceNumber = GetInvoiceNo2("SO033_MediaBillNo", cn)
'                                            .Fields("F97112") = GetString(strSequenceNumber, 16, giLeft, False)
'                                             cn.Execute ("UPDATE SO033 A SET MediaBillNo  ='" & strSequenceNumber & "'  WHERE   BillNo = '" & rsTmp("BillNO") & "'   AND " & strChooseMultiAcc & " AND CustID = " & rsTmp("CUSTID") & " AND AccountNO ='" & rsTmp("AccountNO") & "'")
'                                         End If
'                                    End If
'
                  ''    ''""""""""""""""""""""""""""""""""""""""""""""""""""""     ''""""""""""""""""""""""""""""""""""""""""""""""""""""
                  ''    ''""""""""""""""""""""""""""""""""""""""""""""""""""""     ''""""""""""""""""""""""""""""""""""""""""""""""""""""
                   
                   '' .Fields("MerchantDescription") = rsTmp("BillNO") & Space((20 - Len(rsTmp("BillNO"))))
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
         ''  End If
NextLoop:
        rsTmp.MoveNext
        Wend
        
        
        pSubCount = lngSubCount
        pTotalAmount = lngtotal
        
        BeginTran = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
        '' 203/08/05 �o�@�q���ݭn�F�]���@�w�|�����ХX�檺
'         If lngDoubleMediaNO > 0 Then
'                    Call MsgBox("�X�檺��Ʀ��@�ǬO���Ъ��A�Ьd�\�O���ɽT�{�һݪ���Ƥ��e�A�קK���ХX��  !!", vbExclamation + vbOKOnly, "���ХX��ĵ��")
'                    Call Shell("NotePad " & txtErrPath.Text, vbNormalFocus)
'         End If
  End If
  
Exit Function
ChkErr:
    Call CloseFS(intIndexTxt)
    Call ErrSub(Me.Name, "BeginTran" & lngerror & "-" & strAccountID)
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
              For lngLoopi = 0 To 20
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
              ''  .WriteLine (txtMerchantName.Text) '�H�Υd���b�S���W��
              '' .WriteLine (txtPayDate.Text)      '�дڤ��
                .WriteLine (txtDiscountRate.Text) '�馩�v
                .WriteLine (txtStatement.Text)    '�Ȥ�b�椺�e
                '.WriteLine (txtInVoice.Text)     '�Τ@�s��
                .WriteLine (txtDataPath)         '����ɸ��|
                .WriteLine (txtErrPath.Text)     ' ���D�Ѧ��ɸ��|
        End With
        LogFile.Close
        Set LogFile = Nothing
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Function IsDataOk() As Boolean
Dim strErrMsg As String

  If cboSet.ListIndex = 0 Or cboSet.ListIndex = 1 Then
        MsgBox "���b���� (1.���v 2.�д�) �ثe�L�k�ϥ� !!", vbOKOnly + vbInformation, "�T���@"
        cboSet.SetFocus
        Exit Function
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
  
'  If Not IsNumeric(txtSpcNO.Text) Then
'      MsgBox " �H�Υd���b�S���N�X���ݬ��Ʀr�@!! ", vbOKOnly + vbInformation, "�T���@"
'      txtPayDate.SetFocus
'      Exit Function
'  End If
  
 '' If Len(txtSpcNO.Text & "") = 0 Then strErrMsg = "�H�Υd���b�S���N�X": txtSpcNO.SetFocus: GoTo Warning
 '' If Len(txtMerchantName.Text & "") = 0 Then strErrMsg = "�H�Υd���b�S���W��": txtMerchantName.SetFocus: GoTo Warning
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
    
'checkSameNOSQL = "SELECT SUM(A.ShouldAmt)  ShouldAmt  " & _
'                                        "From " & _
'                                        TableOwnerName & "SO033 A ," & _
'                                        TableOwnerName & "SO001 B ," & _
'                                        TableOwnerName & "SO002A C,  " & _
'                                        TableOwnerName & "SO002 D  " & _
'                                        "Where " & strBankId & " AND  " & _
'                                        "C.CARDNAME  ='" & CardName & "' AND " & _
'                                        "A.CustID  = C.CUSTID AND " & _
'                                        "A.CustID  = D.CUSTID AND " & _
'                                        "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
'                                        "A.UCCode Is Not Null AND " & _
'                                        "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo  AND C.StopFlag = 0  " & _
'                                        IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34sql & "  AND A." & snoType & " = '" & strMNO & "'   " & _
'                                        " Group by A." & snoType & " , C.AccountNO  "
                                
 '' �N C.StopFlag = 0  �o�Ӯ����A�קK�j������
                                        
checkSameNOSQL = "SELECT SUM(A.ShouldAmt)  ShouldAmt  " & _
                                        "From " & _
                                        TableOwnerName & "SO033 A,  " & _
                                        TableOwnerName & "SO002 D  " & _
                                        "Where " & _
                                        "A.CancelFlag = 0 And " & _
                                        "A.CustID  = D.CUSTID AND " & _
                                        "A.SERVICETYPE = D.SERVICETYPE   AND   " & _
                                        "A.UCCode Is Not Null  " & _
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
 If intPara24 = 0 Then
       rsStopYMDateSQL = "SELECT E.StopYM   " & _
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
                "Where " & strBankId & " AND  " & _
                "C.CARDNAME  ='" & CardName & "' AND " & _
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
        GetStopYMDateString = rsStopYMDateString(0)
    Else
        GetStopYMDateString = ""
    End If
    
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "GetStopYMDateString")
End Function
