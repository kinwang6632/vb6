VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSo3312A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�C����b��Ƶn��[SO3312A]"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "So3312A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   2670
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      ButtonCaption   =   "���O����"
   End
   Begin VB.ComboBox cboBillHeadFmt 
      Height          =   300
      Left            =   2250
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   1
      Top             =   630
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6120
      TabIndex        =   9
      Top             =   3495
      Width           =   435
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   8
      Top             =   3525
      Width           =   2925
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1350
      TabIndex        =   10
      Top             =   5460
      Width           =   1410
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
      Height          =   435
      Left            =   4275
      TabIndex        =   11
      Top             =   5490
      Width           =   1410
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   390
      ScaleHeight     =   1335
      ScaleWidth      =   6255
      TabIndex        =   12
      Top             =   3975
      Width           =   6315
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�]�Y�I�ں������ťաA�h�H��I�ں����^"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   720
         TabIndex        =   25
         Top             =   1050
         Width           =   3510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�]�Y���O�H���ťաA�h�H�즬�O�H���^"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   720
         TabIndex        =   16
         Top             =   720
         Width           =   3315
      End
      Begin VB.Label lblAccount 
         AutoSize        =   -1  'True
         Caption         =   "�]�Y���ڤJ�b����ťաA�h�H���b�鬰�J�b��^"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   720
         TabIndex        =   15
         Top             =   420
         Width           =   4095
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   30
         TabIndex        =   14
         Top             =   60
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�]�Y���O�覡�ťաA�h�H�즬�O�覡�^"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   90
         Width           =   3315
      End
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   2250
      TabIndex        =   0
      Top             =   180
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BoxWidth        =   4500
      FldWidth2       =   2500
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilBankCode 
      Height          =   315
      Left            =   2250
      TabIndex        =   3
      Top             =   1050
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   1000
      FldWidth2       =   2250
      F2Corresponding =   -1  'True
   End
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   330
      Top             =   4755
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   2250
      TabIndex        =   2
      Top             =   585
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   2250
      TabIndex        =   4
      Top             =   1455
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth2       =   2500
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdaRealDate 
      Height          =   375
      Left            =   3135
      TabIndex        =   7
      Top             =   3075
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin prjGiList.GiList gilCMCode 
      Height          =   315
      Left            =   2250
      TabIndex        =   5
      Top             =   1860
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth2       =   2500
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilPTCode 
      Height          =   315
      Left            =   2250
      TabIndex        =   6
      Top             =   2250
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth2       =   2500
      F2Corresponding =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "�I�ں���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   480
      TabIndex        =   26
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label lblBillHeadFmt 
      AutoSize        =   -1  'True
      Caption         =   "�h�b�Უ�ͨ̾ڳ]�w"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   480
      TabIndex        =   24
      Top             =   680
      Width           =   1755
   End
   Begin VB.Label lblBankCode 
      AutoSize        =   -1  'True
      Caption         =   "��b�N���Ȧ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   480
      TabIndex        =   23
      Top             =   1080
      Width           =   1170
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
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   480
      TabIndex        =   22
      Top             =   240
      Width           =   585
   End
   Begin VB.Label lblServiceType 
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
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   480
      TabIndex        =   21
      Top             =   645
      Width           =   780
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      Caption         =   "���O�H��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   480
      TabIndex        =   20
      Top             =   1490
      Width           =   780
   End
   Begin VB.Label lblDate1 
      AutoSize        =   -1  'True
      Caption         =   "���ڤJ�b����G"
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
      Left            =   1650
      TabIndex        =   19
      Top             =   3180
      Width           =   1365
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "��ƨӷ��]���|+�ɦW�^�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   750
      TabIndex        =   18
      Top             =   3615
      Width           =   2250
   End
   Begin VB.Label lblCMCode 
      AutoSize        =   -1  'True
      Caption         =   "���O�覡"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Top             =   1920
      Width           =   780
   End
End
Attribute VB_Name = "frmSo3312A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strPrgName As String        '��b�{���W��
Private strEmpNo As String          '���O�H��
Private rsSo3312 As New ADODB.Recordset
Private strFilePath As String       '�O�����|�W��
Private objMediaIn As Object
Private objAction As Object
Dim rsBillHeadFmt As New ADODB.Recordset
Dim intPara24 As Integer
Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        '#7049 ���CD068A.CitemCode By Kin 2015/07/09
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
        'If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If Not GetRS(rsCD019, strCD019SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
        'gimCitemCode.ShowMulti
End Sub
Private Sub cmdCancel_Click()
    On Error Resume Next
        Unload Me
    Exit Sub
End Sub

Private Sub cmdOK_Click()
On Error GoTo chkErr
    Dim intRefNo As Integer
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    intRefNo = 0
    If (Trim(UCase(strPrgName)) = "MEDIAPOST3") Or (Trim(UCase(strPrgName)) = "MEDIAPOST") Then
        intRefNo = GetRsValue("Select Nvl(RefNo,0) From " & GetOwner & "CD018 Where CodeNo=" & gilBankCode.GetCodeNo)
    End If
    Set objMediaIn = CreateObject("MediaIn4.clsInterface")
    With objMediaIn
        .FilePath = strFilePath                          'INI�ɮ׸��|
        .prgName = strPrgName                            '��b�{���W��
        .SourcePath = txtPath.Text                       '��l�ɮ׸��|
        .RealDate = gdaRealDate.GetValue                 '�J�b���
        .UpdEn = garyGi(1)                               '���ʤH��
        .uClctEn = gilClctEn.GetCodeNo & ""              '���O�H��
        .uClctName = gilClctEn.GetDescription & ""
        .uCMCode = gilCMCode.GetCodeNo & ""              '���O�覡
        .uCMName = gilCMCode.GetDescription & ""
        .ServiceType = gilServiceType.GetCodeNo & ""     '�A�����O
        .CompCode = gilCompCode.GetCodeNo & ""
        .BankName = gilBankCode.GetDescription & ""
        .BankId = gilBankCode.GetCodeNo
        '#7049 ��� CD068A.CitemCode By Kin 2015/07/15
        If cboBillHeadFmt.Visible Then
            .uCitemQty = "Select CitemCode From " & GetOwner & "CD068A Where CD068A.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' "
        Else
            .uCitemQty = gimCitemCode.GetQueryCode & ""      '���O����
        End If
        .uGetOwner = GetOwner
        .FlowId = garyGi(17)                             '�y�{����
        .uRefNo = intRefNo
        '#7033 �W�[�I�ں��� By Kin 2015/04/28
        .uPTCode = gilPTCode.GetCodeNo
        .uPTName = gilPTCode.GetDescription
        .uUpdTime = GetDTString(RightNow)                '#4388 �W�[���ʮɶ� By Kin 2009/05/04
        Set .ugcnGi = gcnGi
        If Len(strPrgName & "") = 0 Then
            MsgBox "�]�w�{���W�ٵL�]�εL�ϥ��v���I�I", vbExclamation, "����"
        Else
            Set objAction = .InitObject(strPrgName & "")
            objAction.Action
            Set objAction = Nothing
        End If
    End With
    
    '�Y��s���ѡA�h�����{��
    
    If objMediaIn.UpdState = False Then Exit Sub
    Call DefineRs
    Call ScrToRcd
    Set objMediaIn = Nothing
        
   Exit Sub
chkErr:
   Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub ScrToRcd()
On Error GoTo chkErr
    With rsSo3312
       .AddNew
       .Fields("RealDate") = gdaRealDate.GetValue
       .Fields("FilePath") = txtPath.Text
       If Dir(strFilePath & "\" & strPrgName & "In1.log") <> "" Then
          Kill strFilePath & "\" & strPrgName & "In1.log"
       End If
          .Save strFilePath & "\" & strPrgName & "In1.log"
    End With
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Sub cmdPath_Click()
    On Error GoTo chkErr
        With comdPath
                .FileName = txtPath.Text
                .Filter = "��r�� (*.txt;*.dat;*.all)|*.txt;*.dat;*.all|�Ҧ��ɮ� (*.*)|*.*"
                .InitDir = ReadGICMIS1("ErrLogPath")
                .Action = 1
                txtPath.Text = .FileName
        End With
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdPath_Click")
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo chkErr
    Dim strErrMsg As String
    
        If Len(gilCompCode.GetCodeNo & "") = 0 Then strErrMsg = "���q�O": gilCompCode.SetFocus: GoTo 66
        If intPara24 = 0 And Len(gilServiceType.GetCodeNo & "") = 0 Then strErrMsg = "�A�����O": gilCompCode.SetFocus: GoTo 66
        If gilBankCode.GetCodeNo & "" = "" Then strErrMsg = "��b�N���Ȧ�": gilBankCode.SetFocus: GoTo 66
        If gilClctEn.GetCodeNo = "" Then strErrMsg = "���O�H��": gilClctEn.SetFocus: GoTo 66
        If gilCMCode.GetCodeNo = "" Then strErrMsg = "���O�覡": gilCMCode.SetFocus: GoTo 66
        If InStr(1, gilBankCode.GetDescription, "�ɤs") > 0 Then
           If gdaRealDate.GetValue = "" Then strErrMsg = "�J�b���": gdaRealDate.SetFocus: GoTo 66
        End If
        If Len(txtPath) = 0 Then strErrMsg = "��ƨӷ�": txtPath.SetFocus: GoTo 66
        If Dir(Trim(txtPath.Text)) = "" Then MsgBox "�N����Ƥ��s�b�A�Э��s����I", vbExclamation, "�T���I": Exit Function
        IsDataOk = True
    Exit Function
66:
    Call MsgMustBe(strErrMsg)
  Exit Function
   
chkErr:
    Call ErrSub(Me.Name, "IsdataOK")
End Function

Private Sub Form_Activate()
On Error GoTo chkErr
  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
  
    gilCompCode.SetFocus
    strSQL = "Select PrgName From " & GetOwner & "CD018 Where UPPER(SUBSTR(PrgName,1,5)) = '" & "MEDIA" & "'"
    If rsTmp.State = adStateOpen Then rsTmp.Close
'    rsTmp.CursorLocation = adUseClient
'    rsTmp.Open strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
    If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then Exit Sub
    If rsTmp.EOF Then
        MsgBox "�г]�w�Ȧ�N�X����b�{���W��!", vbExclamation, "ĵ�i"
        Unload Me
    End If
    
    rsTmp.Close
    Set rsTmp = Nothing
  Exit Sub
chkErr:
  ErrSub Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyEscape Then Unload Me
        If KeyCode = vbKeyF2 Then
            If Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "") Then
                gdaRealDate.SetFocus
                Exit Sub
            End If
            If cmdOK.Enabled = True Then cmdOK.Value = True
        End If
End Sub
Private Sub CboAddItem()
    On Error Resume Next
        'If cboBillHeadFmt.ListCount > 0 Then Exit Sub
        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068") Then Exit Sub
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub
Private Sub Form_Load()
    On Error GoTo chkErr
        Screen.MousePointer = vbDefault
        Call subGil
        Call subGim
        Call getDefault
        CboAddItem
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub DefineRs()
On Error GoTo chkErr
  If rsSo3312.State = adStateOpen Then rsSo3312.Close
    With rsSo3312
       .LockType = adLockOptimistic
       .Fields.Append "RealDate", adBSTR, 8
       .Fields.Append "FilePath", adBSTR, 50
       .Open
    End With
  Exit Sub
chkErr:
  ErrSub Me.Name, "DefineRS"
End Sub

Private Sub GetInitData()
    On Error GoTo chkErr
   
        strFilePath = ReadGICMIS1("ErrLogPath")
        If Dir(strFilePath & "\" & GetPrgName & "In1.log") = "" Then
            gdaRealDate.SetValue ""
            txtPath.Text = ""
        Else
            If rsSo3312.State = adStateOpen Then rsSo3312.Close
            rsSo3312.Open strFilePath & "\" & GetPrgName & "In1.log"
            gdaRealDate.SetValue rsSo3312("RealDate").Value & ""
            txtPath.Text = rsSo3312("FilePath").Value & ""
        End If
   Exit Sub
chkErr:
    Call ErrSub(Me.Name, "GetInitData")
End Sub

Private Function GetPrgName() As String
  On Error GoTo chkErr
     Dim rsTmp As New ADODB.Recordset
     Dim strSQL As String

        strSQL = "Select PrgName,EmpNo From " & GetOwner & "CD018 Where CodeNo = " & gilBankCode.GetCodeNo
        Call OpenRecordset(rsTmp, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False)

            If Len(rsTmp("PrgName") & "") = 0 Then
                strPrgName = ""
                MsgBox "�г]�w��b�Ȧ�N�X���{���W��!", vbCritical, "���~"
            Else
                strPrgName = rsTmp("PrgName") & ""
                strEmpNo = rsTmp("EmpNo") & ""
                GetPrgName = strPrgName
            End If
            
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "GetPrgName")
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   If rsSo3312.State = adStateOpen Then
      rsSo3312.Close
      Set rsSo3312 = Nothing
   End If
   Call ReleaseCOM(Me)
End Sub


Private Sub gdaRealDate_Validate(Cancel As Boolean)
    On Error Resume Next
        '#3411 �p�G����J�h���ˬd By Kin 2008/06/02
        If gdaRealDate.GetValue = "" Then Exit Sub
        Cancel = Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "")
End Sub


Private Sub gilBankCode_Change()
  On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
       
       If Len(gilBankCode.GetCodeNo) <> 0 And Len(gilBankCode.GetDescription) <> 0 Then
            GetInitData
            gilClctEn.SetCodeNo strEmpNo
            gilClctEn.Query_Description
            gilCMCode.SetDescription "��b"
            gilCMCode.Query_CodeNo
       End If
       
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gilBankCode_Change")
End Sub

Private Sub subGil()
On Error GoTo chkErr
    SetgiList gilBankCode, "CodeNo", "Description", "CD018", , , , , 3, 24, " Where UPPER(SUBSTR(PrgName,1,5)) = '" & "MEDIA" & "'"
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
    SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 3, 12
    SetgiList gilClctEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20
    SetgiList gilCMCode, "CodeNo", "Description", "CD031", , , , , 3, 12
    '#7033 �W�[�I�ں��� By Kin 2015/04/28
    SetgiList gilPTCode, "CodeNo", "Description", "CD032", , , , , 3, 12, , True
  Exit Sub
chkErr:
  ErrSub Me.Name, "subGil"
End Sub
Private Sub subGim()
    On Error Resume Next
        '���O����
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "�N�X", "�W��", , True

End Sub

Private Sub gilCompCode_Change()
On Error GoTo chkErr
Dim strWhere As String
    garyGi(16) = gilCompCode.GetOwner
    Call subGil
    Call subGim
    If Len(gilCompCode.GetCodeNo & "") = 0 Then gilCompCode.SetFocus: strErrFile = "���q�O": GoTo 99
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    If Len(gilCompCode.GetCodeNo & "") = 0 Then
       strWhere = ""
    Else
       strWhere = " Where CompCode = '" & gilCompCode.GetCodeNo & "'"
    End If
        
'    GiListFilter gilBankCode, , gilCompCode.GetCodeNo
'    gilBankCode.Filter = gilBankCode.Filter & IIf(gilBankCode.Filter = "", " Where ", " And ") & " UPPER(SUBSTR(PrgName,1,5)) = '" & "MEDIA" & "'"
    gilClctEn.Filter = strWhere
  Exit Sub
99:
  MsgMustBe (strErrFile)
  Exit Sub
chkErr:
  ErrSub Me.Name, "gilCompCode_Change"
End Sub

Private Sub gilServiceType_Change()
On Error GoTo chkErr
Dim strWhere As String
           
    If Len(gilServiceType.GetCodeNo & "") = 0 Then
       strWhere = ""
    Else
       strWhere = " Where ServiceType = '" & gilServiceType.GetCodeNo & "' Or ServiceType Is Null"
    End If
    gilCMCode.Filter = strWhere
  Exit Sub
chkErr:
  ErrSub Me.Name, "gilServiceType_Change"
End Sub

Private Sub getDefault()
  On Error GoTo chkErr
    Dim intPara23 As Integer
     gilCompCode.SetCodeNo garyGi(9)
     gilCompCode.Query_Description
        intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
        If intPara24 = 1 Then
            lblBillHeadFmt.Visible = True
            cboBillHeadFmt.Visible = True
            '  �δC��J�b�~Lock
            'If intPara23 = 2 Then gimCitemCode.Enabled = False
            If intPara23 = 2 Then gimCitemCode.IsReadOnly = True
            '#7049 �p�G���h�C��b��A�h�n�N���O�������� By Kin 2015/07/15
            'gimCitemCode.Enabled = False
            gimCitemCode.IsReadOnly = True
            gilServiceType.SetCodeNo ""
            gilServiceType.Query_Description
            gilServiceType.Enabled = False
            gilServiceType.Visible = False
            lblServiceType.Visible = False
        Else
            lblBillHeadFmt.Visible = False
            gimCitemCode.Enabled = True
        End If
     
  Exit Sub
chkErr:
  ErrSub Me.Name, "getDefault"
End Sub
