VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSO1193A 
   Caption         =   "VOD����@�~[SO32E02A]"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   Icon            =   "SO1193A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   5190
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox txtMvodId 
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
      Left            =   2160
      TabIndex        =   3
      Top             =   3540
      Width           =   2055
   End
   Begin VB.CommandButton cmdMvodId 
      Caption         =   "�]��"
      Height          =   315
      Left            =   4290
      TabIndex        =   4
      Top             =   3540
      Width           =   705
   End
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   0
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "...."
      Height          =   315
      Left            =   4560
      TabIndex        =   11
      Top             =   5940
      Width           =   465
   End
   Begin VB.TextBox txtXLSFile 
      Height          =   315
      Left            =   1050
      TabIndex        =   10
      Top             =   5940
      Width           =   3465
   End
   Begin VB.CommandButton cmdSTB 
      Caption         =   "�]��"
      Height          =   315
      Left            =   4290
      TabIndex        =   2
      Top             =   3180
      Width           =   705
   End
   Begin VB.TextBox txtSrNO 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   3210
      Width           =   2055
   End
   Begin VB.TextBox txtCustid 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2160
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. �T�w"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6390
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1875
      TabIndex        =   13
      Top             =   6375
      Width           =   1365
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "�W���έp���G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3615
      TabIndex        =   14
      Top             =   6375
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   60
      TabIndex        =   22
      Top             =   4560
      Width           =   4995
      Begin VB.CheckBox chkForceAmount 
         Caption         =   "���̾ڰѼƪ��B����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2730
         TabIndex        =   37
         Top             =   150
         Width           =   2115
      End
      Begin VB.CheckBox chkRun 
         Caption         =   "���浲��@�~"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   9
         Top             =   840
         Width           =   1635
      End
      Begin VB.CheckBox chkTran2 
         Caption         =   "���ͥ����浲����ӳ���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   8
         Top             =   480
         Width           =   2475
      End
      Begin VB.CheckBox chkTran1 
         Caption         =   "���͹w�p������ӳ���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   7
         Top             =   150
         Value           =   1  '�֨�
         Width           =   2385
      End
   End
   Begin VB.Frame fraLast 
      Caption         =   "�W������O��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   4950
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "����I���:"
         Height          =   180
         Left            =   825
         TabIndex        =   21
         Top             =   300
         Width           =   945
      End
      Begin VB.Label lblTranDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "YY/MM/DD"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1875
         TabIndex        =   20
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "�W�����浲��ɶ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   19
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label lblUpdTime 
         BackColor       =   &H00E0E0E0&
         Caption         =   "YY/MM/DD HH:MM:SS"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1875
         TabIndex        =   18
         Top             =   645
         Width           =   1980
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "�W�����浲��H��: "
         Height          =   180
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label lblUpdEn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "XXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1875
         TabIndex        =   16
         Top             =   975
         Width           =   2760
      End
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   300
      Left            =   1920
      TabIndex        =   33
      Top             =   1440
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   529
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   900
      FldWidth2       =   1855
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   1920
      TabIndex        =   34
      Top             =   1770
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   556
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   900
      FldWidth2       =   1855
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdaEndDate 
      Height          =   270
      Left            =   2160
      TabIndex        =   5
      Top             =   3900
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ReplaceText     =   -1  'True
   End
   Begin Gi_Date.GiDate gdaShouldDate 
      Height          =   270
      Left            =   2160
      TabIndex        =   6
      Top             =   4230
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ReplaceText     =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "MVodID"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   38
      Top             =   3600
      Width           =   675
   End
   Begin VB.Label lblXLSPath 
      Caption         =   "�ɮ׸��|"
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1290
      TabIndex        =   35
      Top             =   4260
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "(�t���餧�e����ƱN�Q����)"
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
      Height          =   390
      Left            =   3480
      TabIndex        =   32
      Top             =   3930
      Width           =   1665
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblEndDate 
      AutoSize        =   -1  'True
      Caption         =   "����������"
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
      Height          =   255
      Left            =   900
      TabIndex        =   31
      Top             =   3930
      Width           =   1200
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCustid 
      AutoSize        =   -1  'True
      Caption         =   "�Ƚs"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   30
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblPara36 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XXXXXXXXX"
      Height          =   255
      Left            =   2160
      TabIndex        =   29
      Top             =   2550
      Width           =   1215
   End
   Begin VB.Label lblPara35 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XXXXXXXXX"
      Height          =   195
      Left            =   2160
      TabIndex        =   28
      Top             =   2190
      Width           =   1200
   End
   Begin VB.Label lbl5 
      AutoSize        =   -1  'True
      Caption         =   "VOD����@�~��ƭ��B"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   2550
      Width           =   1965
   End
   Begin VB.Label lbl4 
      AutoSize        =   -1  'True
      Caption         =   "VOD����@�~���B���B"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   2190
      Width           =   1965
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "���q�O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1260
      TabIndex        =   25
      Top             =   1455
      Width           =   585
   End
   Begin VB.Label lblSrNO 
      AutoSize        =   -1  'True
      Caption         =   "STB�]�Ƭy����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   780
      TabIndex        =   24
      Top             =   3210
      Width           =   1305
   End
   Begin VB.Label lblServicetype 
      AutoSize        =   -1  'True
      Caption         =   "�A�����O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   23
      Top             =   1830
      Width           =   780
   End
End
Attribute VB_Name = "frmSO1193A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#5327 �ݨD By Kin 2009/11/06
Option Explicit
Dim lngTranDate As Long             ' ����I���
Dim rsCD039 As New ADODB.Recordset  ' ���q�O�N�X��
Dim intPara35 As Integer, intPara36 As Integer
Dim strPara35In As String, strPara36In As String
Dim strSQL As String
Dim strChooseString As String
Dim strFormula As String
Dim strCompCode As String
Dim strServiceType As String
Dim strCustId As String
Dim strSrNO As String
Dim strEndDate As String
Dim strMsg As String
Dim blnDO As Boolean
Dim blnGive As Boolean
Private blnTrans As Boolean
Private strCitemCode As String
Private strCitemName As String
Private strRunTime As String
Private blnBeginTrans As Boolean
Private strCMCode As String
Private strCMName As String
Private strUCCode As String
Private strUCName As String
Private strPTCode As String
Private strPTName As String
Private intPara14 As Integer
Private strSalePointCode As String
Private strSalePointName As String
Private rs014 As New ADODB.Recordset
Private blnNOUpdate As Boolean
Private blnNoData As Boolean
Private rs033Clone As ADODB.Recordset
Private rs033VodClone As ADODB.Recordset
Private rs182Clone As ADODB.Recordset
Private strFaciSeqNo As String
Private strFaciSNo As String
Private blnAutoEXE As Boolean
Private blnforceAmount As Boolean
Private strViewXLSName As String
Private strViewName As String
Private strChoose As String
Private blnforceBill As Boolean


Private Sub chkForceAmount_Click()
    blnforceBill = (chkForceAmount.Value = 1)
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub cmdMvodId_Click()
 On Error GoTo ChkErr
    If txtCustid.Text = "" Then
        MsgBox "�п�J�Ȥ�s���A������]�ưʧ@�I", vbInformation, "�T��"
        Exit Sub
    End If
    
    With frmSO1193B
        .uParentForm = Me
        .uEditMode = giEditModeEdit
        .uCustid = txtCustid.Text
        .uServiceType = gilServiceType.GetCodeNo
        .uEditMode = giEditModeEdit
        '#6395 ��]�ƭn�令���� By Kin 2012/12/18
        .uPRCanClick = True
        .uRefNo = 3
        .uFaciSeqNo = txtMvodId.Tag
        .Move (Screen.Width - frmSO1193B.Width) / 2, (Screen.Height - Me.Height) / 2 + cmdMvodId.Top + cmdMvodId.Height + 400
        .Show 1, Me
        txtMvodId.Text = .uMVodId
        txtMvodId.Tag = .uFaciSeqNo
        On Error Resume Next
        SetFocus
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "cmdSTB_Click"
End Sub


Private Sub cmdOpen_Click()
   On Error GoTo ChkErr
        With comdPath
            .DialogTitle = "��ܿ�X���|"
            .Filter = "Microsoft Excel ����ï (*.xls)|*.xls"
            .Action = 1
             If .FileName <> "" Then txtXLSFile = .FileName
        End With
    Exit Sub
ChkErr:
   ErrSub Me.Name, "cmdOpen_Click"
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      Call PreviousRpt(GetPrinterName(5), RptName("SO32E02"), Me.Caption)
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdSTB_Click()
  On Error GoTo ChkErr
    If txtCustid.Text = "" Then
        MsgBox "�п�J�Ȥ�s���A������]�ưʧ@�I", vbInformation, "�T��"
        Exit Sub
    End If
    
    With frmSO1131E
        .uParentForm = Me
        .uEditMode = giEditModeEdit
        .uCustid = txtCustid.Text
        .uServiceType = gilServiceType.GetCodeNo
        .uEditMode = giEditModeEdit
        '#6395 ��]�ƭn�令���� By Kin 2012/12/18
        .uPRCanClick = True
        .uRefNo = 3
        .uDefFaciSeqNo = txtSrNO.Tag
        .Move (Screen.Width - frmSO1131E.Width) / 2, (Screen.Height - Me.Height) / 2 + cmdSTB.Top + cmdSTB.Height + 400
        .Show 1, Me
        txtSrNO.Tag = .uFaciSeqNo
        txtSrNO.Text = .uFaciSNo
        On Error Resume Next
        SetFocus
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "cmdSTB_Click"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = vbKeyF2 Then cmdOk.Value = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    strServiceType = "": strCompCode = "": strEndDate = "": strCustId = "": strSrNO = ""
    strPara35In = "": strPara36In = "": strFaciSeqNo = "": strFaciSNo = ""
    blnGive = False
    SendSQL , , True
    'Call ReleaseCOM(Me)
End Sub
Private Function FillXlsData(ByVal aFile As String, ByRef aMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim rsXls As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Set rsXls = GetXlsRs(aFile, "Sheet1", " * ")
    If rsXls Is Nothing Then Exit Function
    If rsXls.State = adStateClosed Then Exit Function
    If Not GetRS(rsTmp, "SELECT * FROM " & GetOwner & strViewXLSName & " WHERE 1=0 ", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsXls.EOF Then rsXls.MoveFirst
    Do While Not rsXls.EOF
        rsTmp.AddNew
        If rsXls("�Ƚs") & "" = "" Then aMsg = "Excel�t���ťո�ơI���ˬdExcel": Exit Function
        If rsXls("�A�ȱb��") & "" = "" Then aMsg = "Excel�t���ťո�ơI���ˬdExcel": Exit Function
        If rsXls("�]�ƧǸ�") & "" = "" Then aMsg = "Excel�t���ťո�ơI���ˬdExcel": Exit Function
        rsTmp("CUSTID") = rsXls("�Ƚs") & ""
        rsTmp("VODACCOUNTID") = rsXls("�A�ȱb��") & ""
        rsTmp("FACISNO") = rsXls("�]�ƧǸ�")
        'rsTmp("FACISEQNO") = rsXls("FACISEQNO") & ""
        rsTmp.Update
        rsXls.MoveNext
    Loop
    FillXlsData = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "FillXlsData")
End Function
Private Function CrtTmpTable() As Boolean
  On Error GoTo ChkErr
    Dim strString As String
    strViewXLSName = GetTmpViewName
    On Error Resume Next
        strString = "Create Table " & GetOwner & strViewXLSName & "(CUSTID Number(8), VODACCOUNTID Number(15), FACISNO Varchar2(20), FACISEQNO Varchar2(15))"
        gcnGi.Execute strString
        CrtTmpTable = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "CrtTmpTable")
End Function
Private Sub cmdOk_Click()
  On Error GoTo ChkErr
    'If gdaEndDate.GetValue = "" Then MsgBox "�ж�J����������!", vbInformation, "�T��!": Exit Sub
    blnNoData = False
    If Not IsDataOk Then Exit Sub
    strServiceType = gilServiceType.GetCodeNo
    strCompCode = gilCompCode.GetCodeNo
    strEndDate = gdaEndDate.GetValue
    strCustId = txtCustid
    strSrNO = txtSrNO
    strRunTime = RightNow
    If Not subChoose Then Exit Sub
    If Not blnAutoEXE Then cmdCancel.SetFocus
    Me.Enabled = False
    
    
    If chkTran1.Value = 1 Then
        'If Not subChoose(1, strCompCode, strServiceType, strEndDate, , strCustid, strSrNO) Then Exit Sub
        Call subPrint(0)
    End If
    If chkTran2.Value = 1 Then
        'If Not subChoose(2, strCompCode, strServiceType, strEndDate, , strCustid, strSrNO) Then Exit Sub
        Call subPrint(1)
    End If
    If chkRun.Value = 1 Then
        Screen.MousePointer = vbHourglass
        
        If Not blnTrans Then gcnGi.BeginTrans
        If Not CreateBill(blnNoData) Then
            If Not blnTrans Then gcnGi.RollbackTrans
            GoTo 88
        End If
        If blnGive Then
            If Not NewSO062(blnNoData, strMsg) Then
                If Not blnTrans Then gcnGi.RollbackTrans
                GoTo 88
            Else
                blnDO = True
            End If
        Else
            If blnNoData Then
                strMsg = "����@�~�����I"
            Else
                strMsg = "�L�����ơI"
            End If
            blnDO = True
        End If
        If (blnNOUpdate) And (Not blnTrans) Then
            gcnGi.RollbackTrans
        Else
            If Not blnTrans Then gcnGi.CommitTrans
        End If
    End If
    If Not blnAutoEXE Then
        If blnGive = True And chkRun.Value = 1 Then
            Screen.MousePointer = vbDefault
            MsgBox strMsg, vbInformation, "�T��"
            'Unload Me
        Else
            If (Not blnGive) And (chkRun.Value = 1) And (Not blnNOUpdate) Then MsgBox strMsg, vbInformation, "�T��"
        End If
    End If
   
88:
'    If Not blnTrans Then
'        If blnNoData Then
'            gcnGi.RollbackTrans
'        Else
'            gcnGi.CommitTrans
'        End If
'    End If
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    If blnAutoEXE Then Unload Me
    On Error Resume Next
    gcnGi.Execute "Drop Table " & GetOwner & strViewXLSName
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdOk_Click")
End Sub
Private Sub subPrint(ByVal aEXETYPE As Integer)
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strMsg As String
    Dim strPrintSQL As String
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    ReadyGoPrint
    Call subCreateView(aEXETYPE)
    DoEvents
    Set rsTmp = gcnGi.Execute("SELECT Count(*) as intCount From " & strViewName & " Where RowNum=1")
    If rsTmp("intCount") = 0 Then
        If aEXETYPE = 0 Then
            strMsg = "�L�w�p���浲����"
        Else
            strMsg = "�L�����浲����"
        End If
        MsgBox strMsg, vbExclamation, "����"
        SendSQL , , True
    Else
        strPrintSQL = "SELECT * From " & strViewName & " V"
        strFormula = IIf(strFormula = "", "Sort0={V.CustId}", strFormula & ";Sort0={V.CUSTID}")
        strFormula = strFormula & ";Sort1={V.FACISNO};Sort2={V.VODACCOUNTID};Type=" & aEXETYPE + 1
        Call PrintRpt2(GetPrinterName(5), "SO32E2A.RPT", , "�w�p������ӳ��� [SO32E2A]", strPrintSQL, strChooseString, , True, , , strFormula)
        SendSQL , , True
    End If
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
     
  Call ErrSub(Me.Name, "subPrint")
End Sub
'aEXETYPE = 0 �N��w�p���� aEXETYPE =1 �N������
Private Function subCreateView(ByVal aEXETYPE As Integer) As Boolean
    Dim strView As String
    On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
    On Error GoTo ChkErr
    Dim strCrtSQL As String
    strViewName = GetTmpViewName
    '#5521 ���ͥX�Ӫ����B�n�A����SUM3 By Kin 2010/02/03
'    strCrtSQL = "SELECT DISTINCT CUSTID,VODACCOUNTID," & _
'                "FACISNO,SUM(MINUSCREDIT) MINUSCREDIT,SUM(AddCreditit) ADDCREDIT," & _
'                "ABS(SUM(MINUSCREDIT)-SUM(ADDCREDITIT)-SUM3) TOTAL  FROM (" & _
'                strSQL & ") WHERE EXETYPE=" & aEXETYPE & _
'                " GROUP BY CUSTID,VODACCOUNTID,FACISNO,SUM3"
'    strCrtSQL = "SELECT DISTINCT CUSTID,VODACCOUNTID," & _
'                "FACISNO,SUM(MINUSCREDIT) MINUSCREDIT,SUM(AddCreditit) ADDCREDIT," & _
'                "ABS(SUM1-SUM2-SUM3) TOTAL  FROM (" & _
'                strSQL & ") WHERE EXETYPE=" & aEXETYPE & _
'                " GROUP BY CUSTID,VODACCOUNTID,FACISNO,SUM1,SUM2,SUM3"
    
    
    '#6388 SQL�y�k���� By Kin 2013/01/02
    '2013/05/09 Jacky �n�N���b�����������
    strCrtSQL = "SELECT DISTINCT To_Char(CUSTID) CUSTID, TO_CHAR(VODACCOUNTID) VODACCOUNTID ,FACISNO, " & _
            " NOPAYCREDIT MINUSCREDIT,PREPAY ADDCREDIT,(MUSTPAYCREDIT - Nvl(OVERCREDIT,0)+Nvl(hasOverCredit,0) -Nvl(UNPAY,0)) TOTAL," & _
            " (OVERCREDIT+HASOVERCREDIT) OVERCREDIT,SEQNO " & _
            " FROM (" & strSQL & ") WHERE FLAG = " & aEXETYPE
    
    'strSQL = "SELECT * FROM (" & strSQL & ") WHERE EXETYPE=" & aEXETYPE
    strView = "Create View " & strViewName & " as ( " & strCrtSQL & ")"
    gcnGi.Execute strView
    SendSQL strView, True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function
Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gdaEndDate.GetValue & "" = "" Then
        MsgBox "�ж�J����������!", vbInformation, "�T��!"
        If gdaEndDate.Enabled Then gdaEndDate.SetFocus
        Exit Function
    End If
    If chkTran1.Value = 1 Or chkTran2.Value = 1 Or chkRun.Value = 1 Then
        'IsDataOk = True
    Else
        MsgBox "�п�ܭn���檺���ءI", vbInformation, "�T��"
        Exit Function
    End If
    If gilCompCode.GetCodeNo = "" Then
        MsgBox "�п�J���q�O�I", vbInformation, "�T��"
        Exit Function
    End If
    '#5488 �p�G�e������JEXCEL,�h�n�إ߼Ȧs��,�ç�EXCEL��ƶ׶i�� By Kin 2010/01/20
    If txtXLSFile.Text <> "" Then
        If Not CrtTmpTable Then
            MsgBox "�إ�EXCEL�Ȧs��ƪ���~�I", vbInformation, "�T��"
            Exit Function
        End If
        Dim aMsg As String
        If Not FillXlsData(txtXLSFile.Text, aMsg) Then
            If aMsg <> "" Then
                MsgBox aMsg, vbInformation, "�T��"
                Exit Function
            Else
                MsgBox "EXCEL�Ȧs��ƫإߥ��ѡI", vbInformation, "�T��"
                Exit Function
            End If
            
        End If
    End If
    
    IsDataOk = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOK")
End Function
Private Function NewSO062(ByVal aHaveData As Boolean, Optional ByRef aMsg As String) As Boolean
On Error GoTo ChkErr
    Dim strCH As String
    Dim strQry As String
    Dim rs062 As New ADODB.Recordset
    strQry = "SELECT * FROM " & GetOwner & "SO062 " & _
            " WHERE TYPE =4 AND COMPCODE=" & gilCompCode.GetCodeNo & _
            " AND SERVICETYPE = '" & gilServiceType.GetCodeNo & "'"
    If Not GetRS(rs062, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    strCH = "���q�O=" & gilCompCode.GetCodeNo & _
            " ;�A�����O=" & gilServiceType.GetCodeNo & _
            " ; ����������=" & gdaEndDate.GetValue(True) & _
            " ; �Ƚs=" & txtCustid.Text & "; STB�]�ƧǸ�=" & txtSrNO.Text & ""
    If aHaveData Then
        If rs062.EOF Then
            rs062.AddNew
        End If
        With rs062
            .Fields("Type") = 4
            .Fields("TranDate") = gdaEndDate.GetValue(True)
            .Fields("UpdEn") = garyGi(1)
            .Fields("UpdTime") = GetDTString(strRunTime)
            .Fields("Para") = strCH
            .Fields("CompCode") = gilCompCode.GetCodeNo
            .Fields("ServiceType") = gilServiceType.GetCodeNo
            .Update
        End With
        aMsg = "����@�~�����I"
        
    Else
        aMsg = "�L�����ơI"
    End If
    
      

    NewSO062 = True
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "NewSO062")
End Function
Private Function GetSQL2(ByVal aType As Integer, ByVal aWhere As String) As String
    On Error GoTo ChkErr
    Dim strWhere As String
    Dim strWhere2 As String
    Dim strWhere3 As String
    Dim strTotalQry As String
    Dim strSum1Qry As String
    Dim strSum2Qry As String
    Dim strSum3Qry As String
    Dim strSubQry As String
    Dim strRetSQL As String
    Dim strRunSQL As String
    Dim strFields As String
    Dim str004Where As String
    Dim strSubExcel As String
    Dim strExcelWhere As String
    '�,SO004�n�L�o������
    'str004Where = "(A.INSTDATE IS NOT NULL AND A.PRDATE IS NOT NULL AND A.PRDATE > A.INSTDATE)"

    
    
    '****************************�������`�q���I�ƪ� Where ���� *******************************
'    strWhere = " WHERE A.CompCode = B.CompCode And A.ServiceType = B.ServiceType " & _
'                " And CREDITTYPECODE In (SELECT CODENO FROM " & GetOwner & "CD108 " & _
'                " WHERE CD108.REFNO = 3 AND NVL(CD108.STOPFLAG,0) = 0 AND B.SERVICETYPE=CD108.SERVICETYPE)" & _
'                " AND NVL(B.CloseFlag,0)=0 AND NVL(B.STOPFLAG,0)= 0 AND A.SEQNO = B.FACISEQNO "
                
    strWhere = " WHERE A.CompCode = B.CompCode  " & _
                " AND NVL(B.STOPFLAG,0)= 0 AND A.SEQNO = B.FACISEQNO " & _
                " And A.VodAccountid=B.VodAccountId "
    '******************************************************************************************
    
    '****************************�i�ϥ��I�ƪ� Where ���� *****************************************
'    strWhere2 = " WHERE A.CompCode = B.CompCode And A.ServiceType = B.ServiceType " & _
'               " And CREDITTYPECODE In (SELECT CODENO FROM " & GetOwner & "CD108 " & _
'               " WHERE CD108.REFNO IN(1,2) AND NVL(CD108.STOPFLAG,0) = 0 AND B.SERVICETYPE=CD108.SERVICETYPE)" & _
'               " AND NVL(B.STOPFLAG,0)=0 AND A.SEQNO=B.FACISEQNO "
               
    strWhere2 = " WHERE A.CompCode = B.CompCode  " & _
               " AND NVL(B.STOPFLAG,0)=0 AND A.SEQNO=B.FACISEQNO And A.VodAccountid=B.VodAccountId"
    
    '**********************************************************************************************
    strExcelWhere = Empty
    If strViewXLSName <> "" And txtXLSFile.Text <> "" Then
        strExcelWhere = " AND A.FACISNO IN ( SELECT FACISNO FROM " & GetOwner & strViewXLSName & ") " & _
                " AND A.VODACCOUNTID IN ( SELECT VODACCOUNTID FROM " & GetOwner & strViewXLSName & " ) "
    End If
    
    
    '****************************�w����L�������`�q���I�ƪ� Where ���� *******************************
    '�o�@�q�y�k���n�F,�ҥH�[�@��1=0
    '#5432 ��o�q�[�i�h,SUM(SO033.SHOULDAMT) By Kin 2009/12/09
    strWhere3 = " WHERE A.CompCode = B.CompCode  " & _
                " AND NVL(B.STOPFLAG,0)=0 AND A.SEQNO=B.FACISEQNO AND A.VODACCOUNTID=B.VODACCOUNTID "
    '**************************************************************************************************
    If aWhere <> Empty Then
        strWhere = strWhere & " AND " & aWhere
        '#5551 ���M������FSUM1�n��䥦���n�� By Kin 2010/02/22
        If gdaEndDate.GetValue <> "" Then
            Call subAnd2(strWhere, " To_Date(To_Char(B.RecTime,'YYYYMMDD'),'YYYYMMDD') <= " & _
                        "To_Date('" & gdaEndDate.GetValue & "','YYYYMMDD')")
            
        End If
        
        
        strWhere2 = strWhere2 & " AND " & aWhere
        strWhere3 = strWhere3 & " AND " & aWhere
        
    End If
    strWhere2 = strWhere2 & strExcelWhere
    strWhere3 = strWhere3 & strExcelWhere
    '****************************�n���⪺SQL�y�k*****************************************************
    
    '****************************��X�������`�q�ʼƤ����٭n�P�_�X�p�Υ����⦸�ƪ�����*****************
    '#5488 EXCLE����ƪ���,�j���N���⦸�Ƴ]�w���Ѽ�+1�A�H�K�i�H�j���X�b By Kin 2010/01/20
    If strViewXLSName <> "" And txtXLSFile.Text <> "" Then
        strSubExcel = " Union " & _
                "SELECT B.VODACCOUNTID,0 MINUSTOTAL," & intPara36 + 1 & " MAXNOCLOSETIME, " & _
                "0 ADDTOTAL " & _
                " FROM " & GetOwner & " SO004 A," & GetOwner & "SO182 B," & GetOwner & strViewXLSName & " C " & _
                strWhere & " AND A.VODACCOUNTID=C.VODACCOUNTID AND A.CUSTID=C.CUSTID " & _
                " AND NVL(A.FACISNO,'X')=Nvl(C.FACISNO,'X') " & _
                " GROUP BY B.VODACCOUNTID "
    End If
    '�p�G���OEXCEL�����,���n���浲�� By Kin 2010/08/06
    '#6101 �p�G�Ŀ�j���X�b�h�j��X�b By Kin 2011/08/09
    Dim aforceTotal As Long
    aforceTotal = 0
    If blnforceBill Then aforceTotal = 1 + intPara36
    strSubQry = "SELECT B.VODACCOUNTID,SUM(B.MINUSCREDIT) MINUSTOTAL," & _
                " Max(Nvl(B.NOCLOSETIME,0)) + " & aforceTotal & " MAXNOCLOSETIME, " & _
                    "SUM(B.AddCredit) ADDTOTAL " & _
                    " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B" & strWhere & _
                    IIf(strExcelWhere <> "", " AND 1=0 ", " ") & _
                    " GROUP BY B.VODACCOUNTID " & strSubExcel

    strSubQry = "SELECT VODACCOUNTID,SUM(MINUSTOTAL) MINUSTOTAL,SUM(MAXNOCLOSETIME) MAXNOCLOSETIME, " & _
            "SUM(ADDTOTAL) ADDTOTAL FROM (" & strSubQry & ") GROUP BY VODACCOUNTID "
    '**************************************************************************************
    strSum1Qry = "SELECT B.VODACCOUNTID,B.NOCLOSETIME,C.MINUSTOTAL,B.MINUSCREDIT " & _
                        " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B,(" & strSubQry & ") C " & _
                        strWhere & " AND B.VODACCOUNTID=C.VODACCOUNTID "
                        

                         
    ' " AND (B.NOCLOSETIME > " & intPara36 & " OR C.TOTAL > " & intPara35 & ")"
                         
    '��X�������`�q�ʼ�
    strSum1Qry = " SELECT VODACCOUNTID,SUM(MINUSCREDIT) SUM1,0 SUM2,0 SUM3 " & _
                 " FROM (" & strSum1Qry & ") GROUP BY VODACCOUNTID "
                 
    '��X�i�ϥ��I��
    strSum2Qry = "SELECT B.VODACCOUNTID,0 SUM1,SUM(NVL(B.AddCredit,0)) SUM2,0 SUM3 " & _
                " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B " & _
                strWhere2 & " GROUP BY B.VODACCOUNTID "
                
    '***************************************************************************
    '��X�w����L�������`�q���I��
    '#5432 strSum3Qry �hSum SO033.ShouldAmt By Kin 2009/12/09
    strSum3Qry = "SELECT A.CUSTID,B.VODACCOUNTID,A.SEQNO " & _
                " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B" & _
                strWhere3 & " GROUP BY A.CUSTID,B.VODACCOUNTID,A.SEQNO"
                
    strSum3Qry = "SELECT A.VODACCOUNTID,0 SUM1,0 SUM2,SUM(B.SHOULDAMT) SUM3 " & _
                " FROM (" & strSum3Qry & ") A," & GetOwner & "SO033 B," & _
                GetOwner & "CD019 C, " & GetOwner & "CD013 D " & _
                " WHERE B.CUSTID=A.CUSTID AND B.FACISEQNO=A.SEQNO AND B.UCCODE IS NOT NULL " & _
                " AND B.CITEMCODE=C.CODENO AND NVL(C.REFNO,0)=21  " & _
                " AND B.UCCODE=D.CODENO AND NVL(D.REFNO,0) Not In (3,7) " & _
                " AND B.BILLNO || B.ITEM NOT IN(SELECT NVL(B.BILLNO,'X')||NVL(B.ITEM,0) FROM " & _
                GetOwner & "SO004 A," & GetOwner & "SO182 B " & strWhere3 & ") " & _
                " GROUP BY A.VODACCOUNTID "
    '***************************************************************************
    
'    strSum3Qry = "SELECT B.VODACCOUNTID,0 SUM1,0 SUM2,SUM(NVL(B.MinusCredit,0)) SUM3 " & _
'                " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B" & _
'                strWhere3 & " GROUP BY B.VODACCOUNTID "
    
    '�N��ƦX��
    strTotalQry = "SELECT VODACCOUNTID,SUM(SUM1) SUM1, " & _
                " SUM(SUM2) SUM2,SUM(SUM3) SUM3 FROM (" & _
                strSum1Qry & " UNION " & strSum2Qry & " UNION " & strSum3Qry & _
                " ) GROUP BY VODACCOUNTID "
                
    '#5432 Add+SO33.ShouldAmt�n<Minus �~�ݭn���M By Kin 2009/12/10
    strRunSQL = " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B ,(" & strSubQry & ") C, " & _
                "(" & strTotalQry & ") D " & _
                strWhere & " AND B.VODACCOUNTID=C.VODACCOUNTID " & _
                " AND B.VODACCOUNTID=D.VODACCOUNTID " & _
                " AND D.SUM2+SUM3-SUM1 < 0 "
    '�̫�ҭn�D�諸���
    strFields = "SELECT A.CUSTID,A.DeclarantName,B.VodAccountid,B.FaciSeqNo," & _
                " B.NoCloseTime,C.MINUSTOTAL,C.ADDTOTAL,B.OrderNo,NVL(B.MinusCredit,0) MINUSCREDIT, " & _
                " NVL(B.AddCredit,0) ADDCREDITIT,B.RecTime,A.FACISNO,B.ROWID ID,C.MAXNOCLOSETIME,Nvl(D.SUM1,0) SUM1, " & _
                " NVL(D.SUM2,0) SUM2,NVL(D.SUM3,0) SUM3 "

'                " AND ( B.NOCLOSETIME > " & intPara36 & " OR C.TOTAL > " & intPara35 & ")"
    '********************************************************************************************************
    Select Case aType
        Case 1
            strRetSQL = strFields & ",0 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(C.MAXNOCLOSETIME,0) >= " & intPara36 & " OR C.MINUSTOTAL >= " & intPara35 & ")"
        Case 2
            strRetSQL = strFields & ",1 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(C.MAXNOCLOSETIME,0) < " & intPara36 & " AND C.MINUSTOTAL < " & intPara35 & ")"
        Case 3
            '#5521 �p��O�έ쥻�u�P�_SUM1-SUM2,���W��n�h����SUM3�����B By Kin 2010/02/03
'            strRetSQL = strFields & ",0 EXETYPE " & strRunSQL & _
'                        " AND ( Nvl(C.MAXNOCLOSETIME,0) >= " & intPara36 & " OR ABS(Nvl(C.MINUSTOTAL,0)-Nvl(ADDTOTAL,0)-Nvl(SUM3,0)) >= " & intPara35 & ")" & _
'                        " UNION " & _
'                        strFields & ",1 EXETYPE " & strRunSQL & _
'                        " AND ( Nvl(C.MAXNOCLOSETIME,0) < " & intPara36 & " AND ABS(Nvl(C.MINUSTOTAL,0)-Nvl(C.ADDTOTAL,0)-Nvl(SUM3,0)) < " & intPara35 & ")"

            strRetSQL = strFields & ",0 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(C.MAXNOCLOSETIME,0) >= " & intPara36 & " OR ABS(Nvl(SUM1,0)-Nvl(SUM2,0)-Nvl(SUM3,0)) >= " & intPara35 & ")" & _
                        " UNION " & _
                        strFields & ",1 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(C.MAXNOCLOSETIME,0) < " & intPara36 & " AND ABS(Nvl(SUM1,0)-Nvl(SUM2,0)-Nvl(SUM3,0)) < " & intPara35 & ")"

    End Select
    GetSQL2 = strRetSQL
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetSQL2")
End Function


'��������⤧VOD�q�ʸ��
Private Function subChoose() As Boolean
  On Error GoTo ChkErr
    strChoose = Empty
    '************************�e���W������********************************************
    If gilCompCode.GetCodeNo <> "" Then
        Call subAnd2(strChoose, " A.CompCode = " & gilCompCode.GetCodeNo)
    End If
    If gilServiceType.GetCodeNo <> "" Then
        Call subAnd2(strChoose, " A.ServiceType='" & gilServiceType.GetCodeNo & "'")
    End If
    If txtCustid.Text <> "" Then
        Call subAnd2(strChoose, " A.CustId=" & txtCustid.Text)
    End If
    If txtSrNO.Tag <> "" Then
        Call subAnd2(strChoose, " A.SEQNO ='" & txtSrNO.Tag & "'")
    Else
        If txtSrNO.Text <> "" Then
            Call subAnd2(strChoose, "A.FaciSNO='" & txtSrNO.Text & "'")
        End If
    End If
    '#6535 �W�[VODACCOUNTID �P FACISNO ���� By Kin 2013/06/27
    If strViewXLSName <> Empty And txtXLSFile.Text <> "" Then
        Call subAnd2(strChoose, " A.CUSTID || A.VODACCOUNTID || A.FACISNO  " & _
                        " IN ( SELECT CUSTID || VODACCOUNTID || FACISNO " & _
                        " From " & GetOwner & strViewXLSName & ")")
    End If
'    If gdaEndDate.GetValue <> "" Then
'        Call subAnd2(strChoose, " To_Date(To_Char(B.RecTime,'YYYYMMDD'),'YYYYMMDD') <= " & _
'                        "To_Date('" & gdaEndDate.GetValue & "','YYYYMMDD')")
'    End If
   
    '************************************************************************************
     '*************************SO004���]�Ƥ@�w�n�ѦҸ���3�ӥB���঳������********************
    Call subAnd2(strChoose, "A.FaciCode In ( SELECT CODENO FROM " & GetOwner & "CD022 " & _
                        " WHERE NVL(StopFlag,0)=0 AND NVL(REFNO,0)= 3 )")
    '#6395 ����]�n���X,�� A.PRDATE IS NULL ������� By Kin 2012/12/18
    
    If Not blnAutoEXE And 1 = 0 Then
        Call subAnd2(strChoose, "A.PRDATE IS NULL ")
    End If
    '*****************************************************************************************
    'strSQL = GetSQL2(3, strChoose)
    strSQL = GetSQL(strChoose)
    strChooseString = "���q�O�@: " & subSpace(strCompCode) & ";" & _
                      "�A�����O: " & subSpace(strServiceType) & ";" & _
                      "�����������@: " & subSpace(strEndDate) & ";" & _
                      "PPV����@�~���B���B: " & intPara35 & ";" & _
                      "PPV����@�~��ƭ��B: " & intPara36 & _
                      IIf(txtCustid.Text <> "", ";�Ȥ�s��: " & txtCustid.Text, "") & _
                      IIf(txtSrNO.Text <> "", ";STB�]�Ƭy����:" & txtSrNO.Text, "")
                      
    subChoose = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Function
Private Function GetSQL(ByVal strWhere As String)
  On Error GoTo ChkErr
    Dim aSQL1 As String
    Dim aSQL2 As String
    Dim aExcelWhere1 As String
    Dim aExcelWhere2 As String
    Dim aMvodIdWhere As String
    Dim aPresentSQL As String
    Dim aMVodPresent As String
    aExcelWhere1 = Empty
    aExcelWhere2 = Empty
    'MustPayCredit=A.NOPAYCREDIT - A.PREPAY + A.PRESENT(�����I�O�I��=���I�I��-�x���I��+���I :100-80+20)
    
    aMvodIdWhere = " Select MVodId From " & GetOwner & "SO004G H Where A.VodAccountId= H.VodAccountId "
    '#7135
    aPresentSQL = "Select Nvl(Sum(Nvl(AddCredit,0) - Nvl(UseCredit,0)),0) Present From " & GetOwner & "SO182 " & _
                    " Where Nvl(StopFlag,0) = 0 And  MinusType = 2 " & _
                    " And Trunc(Nvl(StartDate,To_Date('19110101','yyyymmdd'))) >= (Select Trunc(Max ( Nvl(incurredDate,To_Date('30001231','yyyymmdd' )))) From " & _
                                                GetOwner & "SO033VOD  Where " & _
                                                " SO182.VODAccountid =  SO033VOD.VODAccountid " & _
                                                " And SO182.VODAccountId = A.VODAccountId)"
    aPresentSQL = "0"
    
    If Len(txtMvodId.Text & "") > 0 Then aMvodIdWhere = aMvodIdWhere & " AND MVODID IN (" & txtMvodId.Text & ")"
    If strViewXLSName <> "" And txtXLSFile.Text <> "" Then
     
           
        aExcelWhere1 = " ( A.FACISNO || A.VODACCOUNTID IN ( SELECT FACISNO || VODACCOUNTID FROM " & GetOwner & strViewXLSName & ") " & _
                              " AND (NVL(A.MUSTPAYCREDIT,0) - NVL(A.UNPAY,0)-NVL(A.OverCredit,0) ) < " & IIf(blnforceBill, "-1", " nvl(B.PARA35,0)") & ")"
                              
         aExcelWhere2 = " ( A.FACISNO || A.VODACCOUNTID NOT IN ( SELECT FACISNO || VODACCOUNTID FROM " & GetOwner & strViewXLSName & ")) "

    End If
'    aSQL1 = "SELECT RANK() OVER (PARTITION BY B.VODACCOUNTID ORDER BY B.SEQNO DESC) RankX," & _
'                " A.NOPAYCREDIT,A.PREPAY,A.PRESENT,B.VODACCOUNTID,B.SERVICETYPE,B.COMPCODE," & _
'                " B.CUSTID,B.SEQNO,B.FACICODE,B.FACISNO,B.RESEQNO" & _
'                " FROM " & GetOwner & "SO004G A," & GetOwner & "SO004 B " & _
'                " WHERE A.VODACCOUNTID = B.VODACCOUNTID AND A.NOPAYCREDIT > A.PREPAY + (A.PRESENT - " & aPresentSQL & ")" & _
'                " AND B.FACICODE IN (SELECT CODENO FROM " & GetOwner & "CD022 WHERE REFNO = 3 AND NVL(STOPFLAG,0) =0 ) "

    aSQL1 = "SELECT RANK() OVER (PARTITION BY B.VODACCOUNTID ORDER BY B.SEQNO DESC) RankX," & _
                " A.NOPAYCREDIT,A.PREPAY,( A.PRESENT - (" & aPresentSQL & ")) PRESENTX,0 PRESENT ,B.VODACCOUNTID,B.SERVICETYPE,B.COMPCODE," & _
                " B.CUSTID,B.SEQNO,B.FACICODE,B.FACISNO,B.RESEQNO" & _
                " FROM " & GetOwner & "SO004G A," & GetOwner & "SO004 B " & _
                " WHERE A.VODACCOUNTID = B.VODACCOUNTID AND A.NOPAYCREDIT > A.PREPAY - (A.PRESENT - (" & aPresentSQL & "))" & _
                " AND B.FACICODE IN (SELECT CODENO FROM " & GetOwner & "CD022 WHERE REFNO = 3 AND NVL(STOPFLAG,0) =0 ) "
      
    
'    aMVodPresent = "Nvl(Sum((Select Nvl(Sum(Nvl(AddCredit,0) - Nvl(UseCredit,0)),0)  From " & GetOwner & "SO182 " & _
'                    " Where Nvl(StopFlag,0) = 0 And  MinusType = 2 " & _
'                    " And Nvl(StartDate,To_Date('29991231','yyyymmdd')) < TRUNC( TO_Date('" & gdaEndDate.GetValue & "','yyyymmdd')) " & _
'                    " And SO182.VODAccountId = G.VODAccountId)),0)"
     aMVodPresent = "Nvl(Sum((Select Nvl(Sum(Nvl(AddCredit,0) - Nvl(UseCredit,0)),0)  From " & GetOwner & "SO182 " & _
                                        " Where Nvl(StopFlag,0) = 0 And  MinusType = 2 " & _
                    " And Trunc(Nvl(StartDate,To_Date('19110101','yyyymmdd'))) >= (Select Trunc( Max ( Nvl(incurredDate,To_Date('30001231','yyyymmdd' )))) From " & _
                                                                                                     GetOwner & "SO033VOD  Where " & _
                                                                                                     " SO182.VODAccountId = SO033VOD.VODAccountId " & _
                                                                                                     " And SO182.VODAccountId = G.VODAccountId))),0)"
    
    aMVodPresent = "0"
  Dim hasOverCredit As String
  hasOverCredit = " ( select Nvl(sum(Nvl(usecredit,0)),0) From " & GetOwner & "so182  " & _
                            " WHERE so182.creditseqno  in (select seqno " & _
                            " From " & GetOwner & "SO033VOD " & _
                            " Where SO182.vodaccountid = SO033vod.vodaccountid " & _
                            " AND SO033VOD.Seqno Is Not Null " & _
                            " AND TRUNC(so033vod.INCURREDDATE)> TRUNC( TO_Date('" & gdaEndDate.GetValue & "','yyyymmdd'))) " & _
                            " And SO182.VodAccountId = A.VodaccountId) hasOverCredit"
  Dim hasOverCredit2 As String
  hasOverCredit2 = " ( select Nvl(sum(Nvl(usecredit,0)),0) From " & GetOwner & "so182  " & _
                            " WHERE so182.creditseqno  in (select seqno " & _
                            " From " & GetOwner & "SO033VOD " & _
                            " Where SO182.vodaccountid = SO033vod.vodaccountid " & _
                            " AND SO033VOD.Seqno Is Not Null " & _
                            " AND SO033VOD.VODACCOUNTID IN (Select VodAccountId From " & GetOwner & "SO004G H " & _
                            " Where H.MvodId = ( ( " & aMvodIdWhere & "))) " & _
                            " AND TRUNC(so033vod.INCURREDDATE)> TRUNC( TO_Date('" & gdaEndDate.GetValue & "','yyyymmdd'))) " & _
                            " And SO182.VODAccountId = A.VodAccountId ) hasOverCredit2"
    'hasOverCredit = "(select 0 from dual ) hasOverCredit"
    'hasOverCredit2 = "(select 0 from dual ) hasOverCredit2"
    'UNPay �N��w�X�b
     'OR C.SALEPOINTCODE IS NOT NULL)
    aSQL1 = "SELECT (A.NOPAYCREDIT - A.PREPAY - A.PRESENT ) MustPayCredit, " & _
    " (SELECT NVL(SUM(C.SHOULDAMT),0) FROM " & GetOwner & "SO033 C " & _
    " WHERE C.UCCODE IS NOT NULL AND NVL(C.CANCELFLAG,0) = 0 " & _
    " AND A.CUSTID = C.CUSTID AND (A.SEQNO = C.FACISEQNO OR A.RESEQNO = C.FACISEQNO ) " & _
    " AND C.UCCODE NOT IN (SELECT CODENO FROM " & GetOwner & "CD013 WHERE REFNO IN (3,7,8) OR PAYOK = 1) " & _
    " AND C.CITEMCODE IN (SELECT CODENO FROM " & GetOwner & "CD019 WHERE REFNO =21)) UNPAY," & _
    " (SELECT Decode(Min(incurredDate),null,sysdate +365,Min(incurredDate)) FROM (" & _
    "  SELECT VODACCOUNTID,INCURREDDATE,SEQNO,VALUE FROM " & GetOwner & "SO033VOD D Where D.CloseBillNo is null " & _
    " MINUS " & _
    " SELECT F.VODACCOUNTID,F.INCURREDDATE,F.SEQNO,VALUE " & _
    " FROM " & GetOwner & "SO182SO033VOD E ," & GetOwner & "SO033VOD F " & _
    " WHERE E.CREDITSEQNO = F.SEQNO ) " & _
    " B WHERE B.VODACCOUNTID IN (SELECT VodAccountId From " & GetOwner & "SO004G H Where MVodId = ( " & _
    aMvodIdWhere & ")) " & _
    " AND VALUE >0 ) MINUSEDATE," & _
    " (SELECT NVL(SUM(VALUE),0)  FROM " & GetOwner & "SO033VOD B WHERE A.VODACCOUNTID = B.VODACCOUNTID " & _
    " AND TRUNC(INCURREDDATE)>TRUNC( TO_Date('" & gdaEndDate.GetValue & "','yyyymmdd'))) OverCredit ," & _
    hasOverCredit & ",  A.*, " & _
    "Nvl((Select Nvl(Sum(G.NOPAYCREDIT),0) -Nvl(SUM(G.PREPAY),0) - (0 - " & aMVodPresent & ") " & _
        " From " & GetOwner & "SO004G G where MVodId = (" & aMvodIdWhere & ") ),0) MustPayCredit2, " & _
    "Nvl((Select Sum(value) From " & GetOwner & "SO033Vod B " & _
        " Where B.VodAccountId in  (Select VodAccountId From " & GetOwner & "SO004G H " & _
         " Where H.MvodId = ( ( " & aMvodIdWhere & "))) " & _
         " AND TRUNC(B.INCURREDDATE)> TRUNC(TO_DATE('" & gdaEndDate.GetValue & "','yyyymmdd'))),0) OverCredit2, " & hasOverCredit2 & _
    " FROM ( " & aSQL1 & " ) "
            
    'sum(Nvl(G.PRESENT,0))
            
    aSQL2 = aSQL1
    
    aSQL1 = "SELECT A.*,0 FLAG,B.Para35,B.Para36 FROM ( " & aSQL1 & " A WHERE A.RANKX=1) A," & GetOwner & "SO043 B " & _
            " WHERE A.SERVICETYPE = B.SERVICETYPE AND (NVL(A.MUSTPAYCREDIT2,0) - NVL(A.UNPAY,0) - NVL(A.OVERCREDIT2,0)+NVL(A.HASOVERCREDIT2,0)  ) > 0 " & _
            " AND (NVL(A.MUSTPAYCREDIT,0) - NVL(A.UNPAY,0) - NVL(A.OVERCREDIT,0) +NVL(A.HASOVERCREDIT,0)) > 0 " & _
            " AND ((nvl(A.MUSTPAYCREDIT2,0) - nvl(A.UNPAY,0)-nvl(A.OverCredit2,0)+Nvl(A.HasOverCredit2,0) ) >=" & IIf(blnforceBill, "1", " nvl(B.PARA35,0) ") & _
            " OR ADD_MONTHS(TO_DATE(TO_CHAR(A.MINUSEDATE,'yyyymm') || '01','yyyymmdd'),NVL(B.PARA36,0)) <=To_Date('" & gdaEndDate.GetValue & "','yyyymmdd') " & _
            IIf(aExcelWhere1 = Empty, "", " OR " & aExcelWhere1) & " )" & _
            " AND " & strWhere
            
    aSQL2 = "SELECT A.*,1 FLAG,B.Para35,B.Para36 FROM ( " & aSQL2 & " A WHERE A.RANKX=1) A," & GetOwner & "SO043 B " & _
            " WHERE A.SERVICETYPE = B.SERVICETYPE AND (NVL(A.MUSTPAYCREDIT2,0) - NVL(A.UNPAY,0) - NVL(A.OVERCREDIT2,0)+NVL(A.HASOVERCREDIT2,0) ) > 0 " & _
            " AND ((nvl(A.MUSTPAYCREDIT2,0) - nvl(A.UNPAY,0)-nvl(A.OverCredit2,0)+nvl(A.HasOverCredit2,0) ) < " & IIf(blnforceBill, "-1", " nvl(B.PARA35,0) ") & _
            " AND ADD_MONTHS(TO_DATE(TO_CHAR(A.MINUSEDATE,'yyyymm') || '01','yyyymmdd'),NVL(B.PARA36,0)) > To_Date('" & gdaEndDate.GetValue & "','yyyymmdd') " & _
            IIf(aExcelWhere2 = Empty, "", " AND " & aExcelWhere2) & " )" & _
            " AND " & strWhere


    GetSQL = aSQL1 & " UNION ALL " & aSQL2
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetSQL")
End Function
Private Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table SO1156"
  On Error GoTo ChkErr
    cnn.Execute "Create Table SO1156 (CompCode text(3),Custid Text(8),AuthenticId text(10),Amount Long,48CodeNo text(10))"
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "CreateTable")
End Sub

Private Function CreateBill(ByRef aHaveData As Boolean) As Boolean
  On Error GoTo ChkErr
    
    Dim rs182 As New ADODB.Recordset
    Dim rsUpd182 As New ADODB.Recordset
    Dim strUpdSql As String
    Dim aVodAccountId As String
    Dim strBillNo As String
    Dim strUPD182 As String
    Dim strUPD033Vod As String
    Dim strSO182SeqNo As String
    Dim strSO033VODSeqno As String
    
    aVodAccountId = Empty

'    If Not GetRS(rs182, "SELECT * FROM (" & strSQL & ") Order By VodAccountId,EXETYPE  ", gcnGi, _
'                adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function

    '#6388 �y�k���� By Kin 2013/01/02
    If Not GetRS(rs182, "SELECT * FROM (" & strSQL & ") Order By VodAccountId,FLAG  ", gcnGi, _
                adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
                
                    
    If rs182.EOF Then
        aHaveData = False
        CreateBill = True
        strMsg = "�L�����ơI"
        Exit Function
    End If
                          
    strUpdSql = "UPDATE " & GetOwner & "SO182 Set NoCloseTime=Nvl(NoCloseTime,0)+1," & _
                "UpdEn='" & garyGi(1) & "',UPDTime='" & GetDTString(strRunTime) & "' "
    rs182.MoveFirst
    
     
    Do While Not rs182.EOF
    
        strSO033VODSeqno = " SELECT SEQNO FROM " & GetOwner & "SO033VOD " & _
                                         " WHERE VODACCOUNTID= " & rs182("VODACCOUNTID") & _
                                         " AND CLOSEBILLNO IS NULL AND VALUE > 0 " & _
                                         " AND TRUNC(INCURREDDATE) <= TO_DATE('" & gdaEndDate.GetValue & "','YYYYMMDD') " & _
                                         " MINUS SELECT A.CREDITSEQNO FROM " & GetOwner & "SO182SO033VOD A," & GetOwner & "SO182 B " & _
                                         " WHERE A.HISCREDITSEQNO=B.SEQNO AND B.VODACCOUNTID= " & rs182("VODACCOUNTID") & _
                                         " AND ( B.CLOSEBILLNO IS NOT NULL Or B.UseCredit >= B.MinusCredit )"
                                         
         strSO182SeqNo = "SELECT SEQNO FROM " & GetOwner & "SO182 " & _
                                         " WHERE VODACCOUNTID = " & rs182("VODACCOUNTID") & _
                                         "  AND NVL(ADDCREDIT,0)=0 AND CLOSEBILLNO IS NULL " & _
                                         " AND CREDITSEQNO IN ( " & strSO033VODSeqno & " ) " & _
                                         " AND Nvl(UseCredit,0) < Nvl(MinusCredit,0) "
        If Val(rs182("Flag")) = 0 Then
            If aVodAccountId <> rs182("VODACCOUNTID") Then
                If Not Insert033(strBillNo, rs182) Then Exit Function
            End If
            
             If blnNOUpdate Or blnAutoEXE Then
                 If Not GetRS(rsUpd182, "SELECT ROWID ID FROM " & GetOwner & "SO182 WHERE SEQNO IN (" & strSO182SeqNo & ")", _
                            gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
             End If
            
                    
            '�N�����⪺���CloseFlag UPD��1,�ö�JBillNO By Kin
            '�p�GBillNo�w�g���Ȫ��ܴN�έ쥻���� By Kin 2009/11/11
            '#5327���դ�OK BillNo�令���n��F By Kin 2009/11/18
            
            
            '#6388 ����strUPD182 By Kin 2013/01/14
            '#�S�n�٭�^�� For Jacky By Kin 2013/05/21
            '#�٭�^�ӭn���ܻy�k
'            strUPD182 = "UPDATE " & GetOwner & "SO182 Set " & _
'                        "CloseTime=To_Date('" & Format(strRunTime, "YYYYMMDDHHMMSS") & "','yyyymmddhh24miss')," & _
'                        "CloseEN='" & garyGi(0) & "',CloseName='" & garyGi(1) & "'," & _
'                        "UpdEn='" & garyGi(1) & "',UPDTime='" & GetDTString(strRunTime) & "', " & _
'                        "CloseBillNo='" & strBillNo & "'," & _
'                        "CloseItem=1,CloseFlag = 1" & _
'                        " Where RowId = '" & rs182("ID") & "' AND NVL(AddCredit,0)= 0 " & _
'                        " AND CLOSEBILLNO IS NULL "
            strUPD182 = "UPDATE " & GetOwner & "SO182 SET " & _
                        "CloseTime=To_Date('" & Format(strRunTime, "YYYYMMDDHHMMSS") & "','yyyymmddhh24miss')," & _
                        "CloseEN='" & garyGi(0) & "',CloseName='" & garyGi(1) & "'," & _
                        "UpdEn='" & garyGi(1) & "',UPDTime='" & GetDTString(strRunTime) & "', " & _
                        "CloseBillNo='" & strBillNo & "'," & _
                        "CloseItem=1,CloseFlag = 1" & _
                        " Where SEQNO IN   ( " & strSO182SeqNo & " )"
                         
            '�HCreditSeqNo��JSO033VOD��CloseBillNo
            '#6388 ����strUPD033Vod By Kin 2013/01/14
            '#�S�n�٭�^�� For Jacky By Kin 2013/05/21
'             strUPD033Vod = " UPDATE " & GetOwner & "SO033VOD SET " & _
'                            " CloseBillNo='" & strBillNo & "'," & _
'                            " CloseItem=1" & _
'                            " WHERE SEQNO=(SELECT CreditSeqNo FROM " & GetOwner & "SO182 " & _
'                                            " WHERE ROWID='" & rs182("ID") & "' AND Nvl(ADDCREDIT,0)= 0 )" & _
'                            " And CloseBillNo IS NULL "

            strUPD033Vod = " UPDATE " & GetOwner & "SO033VOD SET " & _
                            " CloseBillNo='" & strBillNo & "'," & _
                            " CloseItem=1" & _
                            " WHERE SEQNO IN (" & strSO033VODSeqno & " )"
            '#6388 ������sSO182�PSO033VOD By Kin 2013/01/14
            '#�S�n�٭�^�� For Jacky By Kin 2013/05/21
            Call gcnGi.Execute(strUPD182)
            Call gcnGi.Execute(strUPD033Vod)
            
        Else
            '�p�G�O������,�n�N���ƥ[1
            Dim strQry As String
            '#6388 ������sSO182 By Kin 2013/01/14
            '�S�n�٭�^��FOR jacky By Kin 2013/05/23
            'strQry = strUpdSql & " WHERE ROWID='" & rs182("ID") & "'"
            'gcnGi.Execute (strUpdSql & " WHERE ROWID='" & rs182("ID") & "' AND NVL(AddCredit,0)= 0")
            strQry = strUpdSql & " WHERE SEQNO IN ( " & strSO182SeqNo & ")"
            
        End If
        If blnNOUpdate Or blnAutoEXE Then
        '#6388 �����^�Ǥw��s��RecordSet By Kin 2013/01/14
        '�S�n�٭�^�� For Jacky By Kin 2013/05/23
        '#6778 �h�P�_State���A�A�_�h�p�G�Ĥ@�����Flag=0 �|�X�� By Kin 2014/04/25
            If Not rsUpd182 Is Nothing Then
                If rsUpd182.State = adStateOpen Then
                    If rsUpd182.RecordCount > 0 Then
                        rsUpd182.MoveFirst
                        Do While Not rsUpd182.EOF
                            If Not InsCloneRs(rsUpd182("ID") & "", "SO182", strBillNo, rs182Clone) Then Exit Function
                            If Not InsCloneRs(rsUpd182("ID") & "", "SO033VOD", strBillNo, rs033VodClone) Then Exit Function
                            rsUpd182.MoveNext
                        Loop
                    End If
                End If
            End If
            
            
        End If
        aVodAccountId = rs182("VODAccountID")
        rs182.MoveNext
    Loop
  
    CreateBill = True
    On Error Resume Next
    aHaveData = True
    Call CloseRecordset(rs182)
    Call CloseRecordset(rsUpd182)
    Exit Function
ChkErr:
   gcnGi.RollbackTrans
  Call ErrSub(Me.Name, "CreateBill")
End Function
Private Function InsCloneRs(ByVal aRowId As String, ByVal aTableName As String, ByVal strBillNo As String, ByRef rsClone As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    '#5521 ���դ�OK,���ઽ����Where��������,�]�����e�w�gUpd�L�F,�ҥH���strBillNo�N�@�w�O�諸 By Kin 2010/02/04
    If UCase(aTableName) = "SO033VOD" Then
        '#5521 �h�W�[ CloseBillNo Is Null���P�_ By Kin 2010/02/03
        strQry = "SELECT A.ROWID,A.* FROM " & GetOwner & aTableName & " A " & _
                " WHERE A.SEQNO IN (SELECT CREDITSEQNO FROM " & GetOwner & "SO182 " & _
                                    " WHERE ROWID='" & aRowId & "'" & _
                                    " AND CLOSEBILLNO='" & strBillNo & "')"
    Else
        
        strQry = "SELECT A.ROWID,A.* FROM " & GetOwner & aTableName & " A WHERE ROWID='" & aRowId & "'" & _
                " AND A.CLOSEBILLNO ='" & strBillNo & "'"
    End If
    If Not GetRS(rsTmp, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsTmp.EOF Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If rsClone.RecordCount = 0 Then
                rsClone.AddNew
            Else
                rsClone.Find "RowId='" & rsTmp("RowId") & "" & "'", , adSearchForward, 1
                If rsClone.EOF Then
                    rsClone.AddNew
                Else
                    GoTo lNext
                End If
            End If
            For i = 0 To rsTmp.Fields.Count - 1
                If Len(rsTmp.Fields(i).Value & "") > 0 Then
                    rsClone.Fields(rsTmp.Fields(i).Name) = rsTmp.Fields(i).Value & ""
                End If
            Next i
            rsClone.Update
lNext:
            rsTmp.MoveNext
        Loop
    End If
    On Error Resume Next
    Call CloseRecordset(rsTmp)
    InsCloneRs = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "InsCloneRs")
End Function

Private Function Insert033(ByRef aBillNo As String, ByRef rsSO182 As Recordset) As Boolean
  On Error GoTo ChkErr
    Dim rs033 As New ADODB.Recordset
    Dim i As Long
    If Not GetRS(rs033, "SELECT A.RowId,A.* FROM " & GetOwner & "SO033 A WHERE 1=0 ", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not Get014Data(rsSO182("CustId")) Then Exit Function
    With rs033
        .AddNew
        .Fields("CustId") = rsSO182("CustID")
        .Fields("CompCode") = gilCompCode.GetCodeNo
        aBillNo = GetInvoiceNo("T", strServiceType)
        .Fields("BillNo") = aBillNo
        .Fields("Item") = 1
        .Fields("CitemCode").Value = strCitemCode
        .Fields("CitemName").Value = strCitemName
        '#5488 �p�G�e���W���U�������,�h�H�e�����󬰥D
        If gdaShouldDate.GetValue & "" <> "" Then
            .Fields("ShouldDate").Value = gdaShouldDate.GetValue(True)
        Else
            .Fields("ShouldDate").Value = GetADdate(strRunTime)
        End If
        '#5521 ���X���O�έ쥻��SUM1-SUM2,�אּSUM1-(SUM2+SUM3) By Kin 2010/02/03
'        .Fields("OldAmt").Value = Abs(Val(rsSO182.Fields("MINUSTotal") & "") - Val(rsSO182.Fields("ADDTOTAL") & "") - Val(rsSO182.Fields("SUM3") & ""))
'        .Fields("ShouldAmt").Value = Abs(Val(rsSO182.Fields("MINUSTotal") & "") - Val(rsSO182.Fields("ADDTOTAL") & "") - Val(rsSO182.Fields("SUM3") & ""))
        
'        .Fields("OldAmt").Value = Abs(Val(rsSO182.Fields("SUM1") & "") - Val(rsSO182.Fields("SUM2") & "") - Val(rsSO182.Fields("SUM3") & ""))
'        .Fields("ShouldAmt").Value = Abs(Val(rsSO182.Fields("SUM1") & "") - Val(rsSO182.Fields("SUM2") & "") - Val(rsSO182.Fields("SUM3") & ""))
        
        '#6388 �y�k�ܧ�,���]�ܧ� By Kin 2013/01/02
        '2013/05/09 Jacky �n�N���b�����������
        .Fields("OldAmt").Value = rsSO182("MUSTPAYCREDIT") - Val(rsSO182("OverCredit")) - Val(rsSO182("UnPay"))
        .Fields("ShouldAmt").Value = .Fields("OldAmt").Value
        
        .Fields("OldPeriod").Value = 0
        .Fields("RealPeriod").Value = 0
        .Fields("CMCODE").Value = IIf(strCMCode = "", Null, strCMCode)
        .Fields("CMName").Value = IIf(strCMName = "", Null, strCMName)
        .Fields("UCCODE").Value = IIf(strUCCode = "", Null, strUCCode)
        .Fields("UCName").Value = IIf(strUCName = "", Null, strUCName)
        .Fields("PTCODE").Value = 1
        .Fields("PTName").Value = "�{��"
        .Fields("ClassCode") = rs014("ClassCode1") & ""
        '.Fields("FaciSeqNo").Value = rsSO182("FaciSeqNo") & ""
         '#6388 �y�k�ܧ�,���]�ܧ� By Kin 2013/01/02
        .Fields("FaciSeqNo").Value = rsSO182("SEQNO") & ""
        .Fields("FaciSNO").Value = rsSO182("FaciSNO") & ""
        .Fields("ServiceType").Value = gilServiceType.GetCodeNo
        .Fields("Note").Value = "VOD�O�ε��ⲣ�ͪ����O���"
        .Fields("CreateTime").Value = strRunTime
        .Fields("UpdTime").Value = GetDTString(strRunTime)
        .Fields("CreateEn").Value = garyGi(0)
        .Fields("UpdEn").Value = garyGi(1)
        .Fields("AddrNo").Value = rs014("AddrNo") & ""
        .Fields("StrtCode").Value = rs014("StrtCode") & ""
        .Fields("MduId").Value = rs014("MduId") & ""
        .Fields("ServCode").Value = rs014("ServCode") & ""
        .Fields("ClctAreaCode").Value = rs014("ClctAreaCode") & ""
        .Fields("OldClctEn").Value = rs014("ClctEn") & ""
        .Fields("OldClctName").Value = rs014("ClctName") & ""
        .Fields("AreaCode").Value = rs014("AreaCode") & ""
        .Fields("ClctEn").Value = rs014("ClctEn") & ""
        .Fields("ClctName").Value = rs014("ClctName") & ""
        .Fields("SalePointCode").Value = IIf(strSalePointCode = "", Null, strSalePointCode)
        .Fields("SalePointName").Value = IIf(strSalePointName = "", Null, strSalePointName)
        '#5488 �W�[����I��� By Kin 2010/01/21
        .Fields("CloseStopDate").Value = gdaEndDate.GetValue(True)
        .Update
    End With
    If blnNOUpdate Or blnAutoEXE Then
        rs033Clone.AddNew
        For i = 0 To rs033.Fields.Count - 1
            If Len(rs033.Fields(i).Value & "") > 0 Then
                rs033Clone.Fields(rs033.Fields(i).Name).Value = rs033.Fields(i).Value & ""
            End If
        Next i
        rs033Clone.Update
    End If
    Insert033 = True
    On Error Resume Next
    Call CloseRecordset(rs033)
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "Insert033")
End Function
Private Function Get014Data(ByVal aCustId As String) As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    strQry = "SELECT A.*,B.CLASSCODE1 FROM " & GetOwner & "SO014 A," & GetOwner & "SO001 B " & _
            " WHERE B.CUSTID=" & aCustId & " AND " & IIf(intPara14 = 1, "B.ChargeAddrNo", "B.InstAddrNo") & "=A.AddrNo" & _
            " AND B.COMPCODE=" & gilCompCode.GetCodeNo & " AND A.COMPCODE=B.COMPCODE"
    If Not GetRS(rs014, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    Get014Data = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "Get014Data")
End Function






Public Sub subAnd(strCH As String)
  On Error GoTo ChkErr
    If strChoose <> "" Then
        strChoose = strChoose & " And " & strCH
    Else
        strChoose = strCH
    End If
    
    
  Exit Sub
ChkErr:
    Call ErrSub("Utility5", "subAnd")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    strServiceType = "D"
    Call subGil
    Call DefaultVal
    gilServiceType.Query_Description
    If blnNOUpdate Or blnAutoEXE Then
        If Not DefineRs(rs033Clone, "SO033") Then Exit Sub
        If Not DefineRs(rs033VodClone, "SO033Vod") Then Exit Sub
        If Not DefineRs(rs182Clone, "SO182") Then Exit Sub
    End If
    
    If strFaciSeqNo <> "" Then
        txtSrNO.Tag = strFaciSeqNo
    End If
    If strFaciSNo <> "" Then
        txtSrNO.Text = strFaciSNo
    End If
    If strEndDate <> "" Then
        gdaEndDate.SetValue strEndDate
    End If
    
    'If blnAutoEXE And strEndDate <> "" Then cmdOk.Value = True
    
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub
Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
        If blnAutoEXE And strEndDate <> "" Then cmdOk.Value = True
End Sub
Private Sub subGil()
  On Error GoTo ChkErr
    SetgiList gilServiceType, "CodeNo", "Description", "CD046"
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)
    If strCompCode = "" Then gilCompCode.SetCodeNo garyGi(9) Else gilCompCode.SetCodeNo (strCompCode)
    gilCompCode.Query_Description
    If strServiceType = "" Then Call SelectServType(gilCompCode.GetCodeNo, gilServiceType) Else gilServiceType.SetCodeNo (strServiceType)
     gilServiceType.GetDescription
  Exit Sub
ChkErr:
    ErrSub Me.Name, "subGil"
End Sub
Private Sub DefaultVal()
  On Error GoTo ChkErr
    Dim rsSO062 As New ADODB.Recordset
    Dim rsSO043 As New ADODB.Recordset  ' ���O�Ѽ�
    Dim rsCD019 As New ADODB.Recordset
    Dim rs044 As New ADODB.Recordset
    Dim rs107 As New ADODB.Recordset
    Dim strSQL As String, strSO043 As String
    Dim strCD019 As String, strSO044 As String
    Dim strCD017 As String
    'If strServiceType <> "" Then gilServiceType.SetCodeNo strServiceType
    gilCompCode.SetCodeNo gCompCode
    
    
    strSQL = "SELECT * FROM " & GetOwner & "SO062 WHERE TYPE = 4 And CompCode=" & gilCompCode.GetCodeNo & " And ServiceType='" & gilServiceType.GetCodeNo & "'"
    If Not GetRS(rsSO062, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then
        ErrHandle "�}�� [�̪�鵲�ѼưO����](SO062) �ɵo�Ϳ��~: " & Err.Description & " �o�ӵ{���Y�N�����C", False, 2, "Form_Load"
        Unload Me
        Exit Sub
    End If
    
    
    '***************************************************************************************
    '���O�ѼƸ��
    strSO043 = "Select para35,para36,Para14 from " & GetOwner & "SO043 Where CompCode=" & gilCompCode.GetCodeNo & " And ServiceType='" & gilServiceType.GetCodeNo & "'"
    If Not GetRS(rsSO043, strSO043, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Sub
    
    If Not rsSO043.EOF Then
        intPara14 = Val(rsSO043("Para14") & "")
    Else
        intPara14 = 0
    End If
    '***************************************************************************************
    
    '***************************************************************************
    '�s�W��SO033�����O����
    strCD019 = "SELECT CodeNo,Description FROM " & GetOwner & "CD019 " & _
                " WHERE NVL(STOPFLAG,0)=0 AND REFNO = 21 " & _
                " AND SERVICETYPE ='" & gilServiceType.GetCodeNo & "'" & _
                " AND SIGN = '+'"
    If Not GetRS(rsCD019, strCD019, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not rsCD019.EOF Then
        strCitemCode = rsCD019("CODENO") & ""
        strCitemName = rsCD019("DESCRIPTION") & ""
    Else
        strCitemCode = ""
        strCitemName = ""
    End If
    '*****************************************************************************
    
    '***********************SO044���w�]��****************************************
    strSO044 = "SELECT A.CMCode,B.Description CMNAME FROM " & GetOwner & "SO044 A," & GetOwner & "CD031 B " & _
                " WHERE A.CMCODE = B.CODENO AND A.SERVICETYPE='" & gilServiceType.GetCodeNo & "'"
    If Not GetRS(rs044, strSO044, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not rs044.EOF Then
        strCMCode = rs044("CMCODE") & ""
        strCMName = rs044("CMNAME") & ""
    Else
        strCMCode = "": strCMName = ""
    End If
    
    strSO044 = "SELECT A.UCCODE,B.Description UCNAME FROM " & GetOwner & "SO044 A," & GetOwner & "CD013 B " & _
                " WHERE A.UCCODE = B.CODENO AND A.SERVICETYPE='" & gilServiceType.GetCodeNo & "'"
    If Not GetRS(rs044, strSO044, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not rs044.EOF Then
        strUCCode = rs044("UCCODE") & ""
        strUCName = rs044("UCNAME") & ""
    Else
        strUCCode = "": strUCName = ""
    End If
    strPTCode = "1"
    strPTName = "�{��"
    
    '****************************************************************************
    '****************************CD107���w�]��***********************************
    strCD017 = "SELECT CodeNo,Description FROM " & GetOwner & "CD107 " & _
            " WHERE NVL(STOPFLAG,0)=0 AND DefaultValue=1 AND COMPCODE=" & gilCompCode.GetCodeNo
    
    If Not GetRS(rs107, strCD017, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not rs107.EOF Then
        strSalePointCode = rs107("CODENO") & ""
        strSalePointName = rs107("DESCRIPTION") & ""
    Else
        strSalePointCode = ""
        strSalePointName = ""
    End If
    
    '****************************************************************************
    
    
    ' �YSO062�L��� �εL�ǤJ��
    If rsSO062.EOF Then
        ' �h�U��쬰�ťթΪ��
        lblTranDate.Caption = ""
        lblUpdTime.Caption = ""
        lblUpdEn.Caption = ""
        lblPara35.Caption = "0"
        lblPara36.Caption = "0"
        gdaEndDate.SetValue ""
        If Not rsSO043.EOF Then
            intPara35 = rsSO043("Para35").Value & ""
            lblPara35.Caption = rsSO043("Para35").Value & ""
            intPara36 = rsSO043("Para36").Value & ""
            lblPara36.Caption = rsSO043("Para36").Value & ""
        End If
    ' �Y�����
    Else
        '����I���
        'lblTranDate.Caption = Format(rsSO062("TranDate").Value & "", "ee/MM/DD")
        lblTranDate.Caption = GetDTString(rsSO062("TranDate").Value & "")
        lngTranDate = Val(rsSO062("TranDate").Value & "")
        '�W�����浲��ɶ�
        lblUpdTime.Caption = rsSO062("UpdTime").Value & ""
        '�W�����浲��H��
        lblUpdEn.Caption = rsSO062("UpdEn").Value & ""
        'PPV����@�~���B���B PPV����@�~��ƭ��B
        intPara35 = rsSO043("Para35").Value & ""
        lblPara35.Caption = rsSO043("Para35").Value & ""
        intPara36 = rsSO043("Para36").Value & ""
        lblPara36.Caption = rsSO043("Para36").Value & ""
        '���������W���������[1��
        '#5343 �n��w�]�Ȯ��� By Kin 2009/11/27
        'gdaEndDate.SetValue DateAdd("d", 1, rsSO062("TranDate").Value & "")
        
    End If
    
    '�P�_�O�_�ѧO��Form�i�Ӫ�
    txtCustid.Text = strCustId
    If strCustId <> "" Then txtCustid.Enabled = False
    txtSrNO.Text = strSrNO
    txtXLSFile.Visible = False
    cmdOpen.Visible = False
    lblXLSPath.Visible = False
    If blnGive Then
        txtCustid.Enabled = False
        txtSrNO.Enabled = False
        cmdSTB.Enabled = False
        txtMvodId.Enabled = False
        cmdMvodId.Enabled = False
        txtXLSFile.Visible = True
        lblXLSPath.Visible = True
        cmdOpen.Visible = True
        If strEndDate <> "" Then gdaEndDate.SetValue strEndDate & ""
        If strEndDate <> "" Then gdaEndDate.Enabled = False
        If strPara35In <> "" Then lblPara35.Caption = strPara35In Else lblPara35.Caption = intPara35
        If strPara36In <> "" Then lblPara36.Caption = strPara36In Else lblPara36.Caption = intPara36
        chkTran1.Value = 1
        chkTran2.Value = 1
        chkRun.Value = 1
        If strCompCode <> "" Then gilCompCode.Enabled = False
        If strServiceType <> "" Then gilServiceType.Enabled = False
    End If
    If blnAutoEXE Then
        chkTran1.Value = 0
        chkTran2.Value = 0
        chkRun.Value = 1
    End If
    If strFaciSNo <> "" Then
        txtSrNO.Text = strFaciSNo
    End If
    If strFaciSeqNo <> "" Then
        txtSrNO.Tag = strFaciSeqNo
    End If
    If strEndDate <> "" Then gdaEndDate.SetValue strEndDate & ""
    If strEndDate <> "" Then gdaEndDate.Enabled = False
    If strPara35In <> "" Then
        intPara35 = Val(strPara35In)
        lblPara35.Caption = intPara35
    End If
    If strPara36In <> "" Then
        intPara36 = Val(strPara36In)
        lblPara36.Caption = intPara36
    End If
    '#5343 ���դ�OK�A���浲��n���v������ By Kin 2009/11/25
    If Not blnAutoEXE Then
        chkRun.Enabled = GetUserPriv("SO32E02", "(SO32E021)")
    End If
    If Not chkRun.Enabled Then
        chkRun.Value = 0
    End If
    chkForceAmount.Value = 0
    If blnforceAmount Then
        chkForceAmount.Value = 1
    End If
    
    On Error Resume Next
    Call CloseRecordset(rsSO062)
    Call CloseRecordset(rsCD019)
    Call CloseRecordset(rs107)
    Exit Sub
ChkErr:
   ErrSub Me.Name, "DefaultVal"
  
End Sub

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    DefaultVal
  Exit Sub
ChkErr:
   ErrSub Me.Name, "gilServiceType_Change"
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Dim strSQL As String
    garyGi(16) = gilCompCode.GetOwner
    If Len(gilCompCode.GetCodeNo & "") = 0 Then Exit Sub
    ' ���q�O�G�Y���ȥB���}������
    If rsCD039.State = adStateOpen Then rsCD039.Close
    strSQL = "SELECT VouPath,CODENO,EmcCompCode,AccountVer FROM " & GetOwner & "CD039 Where CODENO = " & gilCompCode.GetCodeNo & " ORDER BY CODENO"
    
    If Not GetRS(rsCD039, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then
        ErrHandle "�b�}�� [���q�O�N�X��](CD039) �ɵo�Ϳ��~: " & Err.Description & " �o�ӵ{���Y�N����!", False, 2, "gilCompCode_Validate"
        Exit Sub
    End If
    
     'Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
     DefaultVal
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gilCompCode_Change"
End Sub

Private Sub gilCompCode_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then MsgMustBe ("���q�O"): Cancel = True
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub
Private Function DefineRs(ByRef rsClone As Recordset, ByVal sTableName As String) As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    
    strQry = "SELECT A.ROWID,A.* FROM " & GetOwner & sTableName & " A WHERE 1=0"
    If Not GetRS(rsTmp, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    With rsClone
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        For i = 0 To rsTmp.Fields.Count - 1
            .Fields.Append rsTmp.Fields(i).Name, adBSTR, rsTmp.Fields(i).DefinedSize, adFldIsNullable
        Next i
    End With
    rsClone.Open
    
    DefineRs = True
    On Error Resume Next
    Call CloseRecordset(rsTmp)
    Exit Function
ChkErr:
  ErrSub Me.Name, "DefineRs"
End Function
Public Property Let uFromBat(ByVal vData As Boolean)
    blnGive = vData
End Property
Public Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
    'blnGive = True
End Property

Public Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
    strServiceType = "D"    '#5327 ���դ�OK,�n�T�wD�A��
    'gilServiceType.SetCodeNo vData
    'blnGive = True
End Property

Public Property Let uCustid(ByVal vData As String)
    strCustId = vData
    'txtCustid.Text = vData
    'blnGive = True
End Property

'STB�]�ƧǸ�
Public Property Let uSrNO(ByVal vData As String)
    strSrNO = vData
    txtSrNO.Text = vData
    'txtSrNO.Text = vData
End Property
Public Property Let uFaciSeqNo(ByVal vData As String)
  On Error Resume Next
    strFaciSeqNo = vData
    'txtSrNO.Tag = strFaciSeqNo
End Property
Public Property Let uFaciSNo(ByVal vData As String)
  On Error Resume Next
    strFaciSNo = vData
    'txtSrNO.Text = strFaciSNO
End Property

'����������
Public Property Let uEndDate(ByVal vData As String)
    strEndDate = vData
End Property

'PPV����@�~���B���B
Public Property Let uPara35(ByVal vData As Integer)
    strPara35In = vData
End Property

'PPV����@�~��ƭ��B
Public Property Let uPara36(ByVal vData As Integer)
    strPara36In = vData
End Property
'�^�ǰ��檺Msg
Public Property Get uMsg() As Variant
    uMsg = strMsg
End Property
'�^�ǩҦ��s�WSO033�����
Public Property Get uRs033() As ADODB.Recordset
  On Error Resume Next
    Set uRs033 = rs033Clone
End Property
'�^�ǩҦ���sSO033VOD�����
Public Property Get uRs033Vod() As ADODB.Recordset
  On Error Resume Next
    Set uRs033Vod = rs033VodClone
End Property
'�^�ǩҦ���sSO182�����
Public Property Get uRs182() As ADODB.Recordset
  On Error Resume Next
    Set uRs182 = rs182Clone
End Property
Public Property Set uRs182(ByRef rs As ADODB.Recordset)
    Set rs182Clone = rs
End Property
Public Property Set uRs033(ByRef rs As ADODB.Recordset)
    Set rs033Clone = rs
End Property
Public Property Set uRs033Vod(ByRef rs As ADODB.Recordset)
    Set rs033VodClone = rs
End Property

'��Miggie���M�L�Ӫ�,�����\������Update,��̫�nRollbackTrans
Public Property Let uNoUpdate(ByVal vData As Boolean)
  On Error Resume Next
    blnNOUpdate = vData
End Property
'�O�_���`���槹�{��
Public Property Get uDo() As Boolean
    uDo = blnDO
End Property
'�O�_�����檺���
Public Property Get uHaveData() As Boolean
  On Error Resume Next
    uHaveData = blnNoData
End Property
Public Property Let uAutoExe(ByVal vData As Boolean)
  On Error Resume Next
    blnAutoEXE = vData
End Property

Public Property Let uTrans(ByVal vData As Boolean)
  On Error Resume Next
    blnTrans = vData
End Property
Public Property Let uforceAmount(ByVal vData As Boolean)
  On Error Resume Next
    blnforceAmount = vData
End Property

Private Sub txtSrNO_KeyDown(KeyCode As Integer, Shift As Integer)
    txtSrNO.Tag = ""
End Sub

Private Sub txtSrNO_KeyPress(keyAscii As Integer)
  On Error Resume Next
    txtSrNO.Tag = ""
End Sub


Private Function GetADdate(DateStr As String) As String
  On Error GoTo ChkErr
    If IsNull(DateStr) Or Len(DateStr) = 0 Then
        GetADdate = ""
    Else
        GetADdate = Format(DateStr, "YYYY/MM/DD")
    End If
  Exit Function
ChkErr:
    Call ErrSub("Sys_Lib", "Function GetADdate")
End Function
