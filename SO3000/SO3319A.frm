VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSO3319A 
   Caption         =   "ACH��b���ڸ�Ƶn�� [SO3319A]"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3319A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '���ݵ�������
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   375
      Left            =   90
      TabIndex        =   23
      Top             =   2130
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      ButtonCaption   =   "���O����"
      IsReadOnly      =   -1  'True
   End
   Begin VB.CheckBox chkZero 
      Caption         =   "���O���ت��B<=0�O�_����"
      Height          =   435
      Left            =   5310
      TabIndex        =   22
      Top             =   870
      Value           =   1  '�֨�
      Width           =   1905
   End
   Begin VB.ComboBox cboBillHeadFmt 
      Height          =   315
      Left            =   1950
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   1
      Top             =   495
      Width           =   3285
   End
   Begin VB.PictureBox Picture1 
      Height          =   1125
      Left            =   330
      ScaleHeight     =   1065
      ScaleWidth      =   6255
      TabIndex        =   15
      Top             =   3480
      Width           =   6315
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�]�Y���O�覡�ťաA�h�H�즬�O�覡�^"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   150
         Width           =   3315
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
         TabIndex        =   18
         Top             =   60
         Width           =   540
      End
      Begin VB.Label lblAccount 
         AutoSize        =   -1  'True
         Caption         =   "�]�Y���ڤJ�b����ťաA�h�H���b�鬰�J�b��^"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   720
         TabIndex        =   17
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�]�Y���O�H���ťաA�h�H�즬�O�H���^"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   720
         TabIndex        =   16
         Top             =   810
         Width           =   3315
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   435
      Left            =   4215
      TabIndex        =   9
      Top             =   4800
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Height          =   465
      Left            =   1290
      TabIndex        =   8
      Top             =   4800
      Width           =   1410
   End
   Begin VB.TextBox txtPath 
      Height          =   360
      Left            =   3060
      TabIndex        =   6
      Top             =   3000
      Width           =   2925
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
      Left            =   6060
      TabIndex        =   7
      Top             =   3000
      Width           =   435
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1935
      TabIndex        =   0
      Top             =   90
      Width           =   3270
      _ExtentX        =   5768
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
      FldWidth2       =   2200
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilBankCode 
      Height          =   315
      Left            =   1935
      TabIndex        =   2
      Top             =   915
      Width           =   3270
      _ExtentX        =   5768
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
      FldWidth2       =   2200
      F2Corresponding =   -1  'True
   End
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   270
      Top             =   4770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   1935
      TabIndex        =   3
      Top             =   1320
      Width           =   3255
      _ExtentX        =   5741
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
      FldWidth2       =   1930
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdaRealDate 
      Height          =   375
      Left            =   2235
      TabIndex        =   5
      Top             =   2540
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
      Left            =   1935
      TabIndex        =   4
      Top             =   1725
      Width           =   3270
      _ExtentX        =   5768
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
      FldWidth2       =   2200
      F2Corresponding =   -1  'True
   End
   Begin VB.Label lblBillHeadFmt 
      AutoSize        =   -1  'True
      Caption         =   "�h�b�Უ�ͨ̾ڳ]�w"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   555
      Width           =   1755
   End
   Begin VB.Label lblCMCode 
      AutoSize        =   -1  'True
      Caption         =   "���O�覡"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   780
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "��ƨӷ��]���|+�ɦW�^�G"
      Height          =   195
      Left            =   690
      TabIndex        =   14
      Top             =   3120
      Width           =   2250
   End
   Begin VB.Label lblDate1 
      AutoSize        =   -1  'True
      Caption         =   "���ڤJ�b����G"
      Height          =   195
      Left            =   690
      TabIndex        =   13
      Top             =   2760
      Width           =   1365
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      Caption         =   "���O�H��"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1380
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "���q�O"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   150
      Width           =   585
   End
   Begin VB.Label lblBankCode 
      AutoSize        =   -1  'True
      Caption         =   "��b�N���Ȧ�"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   975
      Width           =   1170
   End
End
Attribute VB_Name = "frmSO3319A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strPrgName As String        '��b�{���W��
Private rsSo3319 As New ADODB.Recordset
Private strFilePath As String       '�O�����|�W��
Private objAgencyIn As Object
Private objAction As Object
Private strChoose As String
Dim rsBillHeadFmt As New ADODB.Recordset
Dim intPara24 As Integer

Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        '#7049 ���CD068A.CitemCode By Kin 2015/07/20
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
        If Not GetRS(rsCD019, strCD019SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        'If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
End Sub
Private Sub CboAddItem()
    On Error Resume Next
        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068 WHERE ACHTNo IS NOT NULL") Then Exit Sub
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub


Private Sub cmdCancel_Click()
    On Error Resume Next
        Unload Me
    Exit Sub
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim objAction As Object
    Dim objInterface As Object
        If IsDataOk = False Then Exit Sub
        Set objInterface = New clsInterface
        Screen.MousePointer = vbHourglass
        With objInterface
'            .uBeOne = IIf(chkBeOne.Value = 1, True, False)
            .uCompCode = gilCompCode.GetCodeNo
'            .uServiceType = gilServiceType.GetCodeNo
            .uErrPath = ReadGICMIS1("ErrLogPath")
            Set .ugcnGi = gcnGi
            .uBankSQL = " = '" & gilBankCode.GetCodeNo & "'"
            .uPTCode = ""
            .uPTName = ""
            .uUpdEn = garyGi(0)
            .uUpdName = garyGi(1)
            .uCMCode = gilCMCode.GetCodeNo
            .uCMName = gilCMCode.GetDescription
            .uClctEn = gilClctEn.GetCodeNo
            .uClctName = gilClctEn.GetDescription
            .uRealDate = IIf(gdaRealDate.GetValue <> "", gdaRealDate.GetValue, Format(RightNow, "YYYYMMDD"))
            '.uCitemQty = gimCitemCode.GetQueryCode
            .uCitemQty = cboBillHeadFmt.Text
            .uFilePath = txtPath.Text
            .uGetOwner = GetOwner
            .FlowId = garyGi(17)
            .uCkZero = IIf(chkZero.Value = 1, True, False) '�t�����B�O�_����
            Set objAction = New csACHTranReferIn
            Call objAction.Action
        End With
        Set objAction = Nothing
        Set objInterface = Nothing
        Call ScrToRcd
        Screen.MousePointer = vbDefault
        Unload Me
    Exit Sub
ChkErr:
    If Err.Number = 20 Then Resume Next
    If Err.Number = 91 Then
       MsgBox "��b�{���W�ٿ��~�Ϊ̸ӻȦ�S����b��Ʋ��ͥ\��!", vbExclamation, "ĵ�i..."
       Exit Sub
    End If
    Call ErrSub(Me.Name, "cmdOk_Click")
End Sub

Private Sub ScrToRcd()
    On Error GoTo ChkErr
        With rsSo3319
            If .EOF Or .BOF Then
                .AddNew
            End If
            .Fields("RealDate") = gdaRealDate.GetValue
            .Fields("FilePath") = txtPath.Text
            If Dir(GetLogFileName(strFilePath)) <> "" Then
                Kill GetLogFileName(strFilePath)
            End If
            .Save GetLogFileName(strFilePath)
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Sub cmdPath_Click()
    On Error GoTo ChkErr
        With comdPath
                .FileName = txtPath.Text
                .Filter = "��r�� (*.txt)|*.txt|�Ҧ��ɮ� (*.*)|*.*"
                .InitDir = ReadGICMIS1("ErrLogPath")
                .Action = 1
                txtPath.Text = .FileName
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPath_Click")
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo ChkErr
        If Not MustExist(gilCompCode, 2, "���q�O") Then Exit Function
        If Not MustExist(cboBillHeadFmt, , "�h�b�Უ�ͨ̾ڳ]�w") Then Exit Function
        If Not MustExist(gilBankCode, 2, "��b�N���Ȧ�") Then Exit Function
        If Not MustExist(txtPath, , "��ƨӷ�") Then Exit Function
        '���ڤJ�b������ର���Ӥ��
        If gdaRealDate.GetValue <> "" Then
            If CDate(gdaRealDate.GetValue(True)) > RightDate Then
                MsgBox "���ڤJ�b������ର���Ӥ��", , "���~�T��"
                gdaRealDate.SetValue ""
                gdaRealDate.SetFocus
                Exit Function
            End If
        End If
        If Not ChkDirExist(Trim(txtPath.Text)) Then MsgBox "�N����Ƥ��s�b�A�Э��s����I", vbExclamation, "�T���I": Exit Function
        
        IsDataOk = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsdataOK")
End Function


Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyEscape Then Unload Me
        If KeyCode = vbKeyF2 Then
            If gdaRealDate.GetValue & "" <> "" Then
                If Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "") Then
                    gdaRealDate.SetFocus
                    Exit Sub
                End If
            End If
            If cmdOK.Enabled = True Then cmdOK.Value = True
            
        End If
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
        Call DefineRS
        Call subGil
        Call subGim
        Call DefaultValue
        CboAddItem
        Call OpenData
        
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Function DefaultValue()
    On Error Resume Next
    Dim intPara23 As Integer
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        Call GetInitData
End Function

Private Function OpenData() As Boolean
    On Error GoTo ChkErr
    Exit Function
ChkErr:
    ErrSub Me.Name, "OpenData"
End Function

Private Sub DefineRS()
On Error GoTo ChkErr
  If rsSo3319.State = adStateOpen Then rsSo3319.Close
    With rsSo3319
       .LockType = adLockOptimistic
       .Fields.Append "RealDate", adBSTR, 8
       .Fields.Append "FilePath", adBSTR, 50
       .Open
    End With
  Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRS"
End Sub

Private Sub GetInitData()
    On Error GoTo ChkErr
        strFilePath = ReadGICMIS1("ErrLogPath")
        If Trim(GetLogFileName(strFilePath)) = "" Then Exit Sub
        If Not ChkDirExist(GetLogFileName(strFilePath)) Then
            gdaRealDate.SetValue ""
            txtPath.Text = ""
        Else
            If rsSo3319.State = adStateOpen Then rsSo3319.Close
            rsSo3319.Open GetLogFileName(strFilePath)
            gdaRealDate.SetValue rsSo3319("RealDate").Value & ""
            txtPath.Text = rsSo3319("FilePath").Value & ""
        End If
   Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "GetInitData")
End Sub

Private Function GetPrgName() As String
  On Error GoTo ChkErr
     Dim rsTmp As New ADODB.Recordset
     Dim strSQL As String
        strSQL = "Select PrgName From " & GetOwner & "CD018 Where CodeNo = " & gilBankCode.GetCodeNo
        If Not GetRS(rsTmp, strSQL, gcnGi) Then Exit Function
            If Len(rsTmp("PrgName") & "") = 0 Then
                strPrgName = ""
                MsgBox "�г]�w��b�Ȧ�N�X���{���W��!", vbCritical, "���~"
            Else
                strPrgName = rsTmp("PrgName")
                GetPrgName = strPrgName
            End If
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "GetPrgName")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Call CloseRecordset(rsBillHeadFmt)
        Call CloseRecordset(rsSo3319)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(frmSO3319A)
End Sub

Private Sub gdaRealDate_Validate(Cancel As Boolean)
    On Error Resume Next
        '#3411 �p�G����J�h���ˬd By Kin 2008/06/02
        If gdaRealDate.GetValue = "" Then Exit Sub
        Cancel = Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "")
End Sub

'�p�G���b�鬰�ťչw�]����Ѥ��
'Private Sub gdaRealDate_GotFocus()
'  On Error Resume Next
'    If gdaRealDate.GetValue <> "" Then Exit Sub
'    gdaRealDate.SetValue Date
'
'End Sub
Public Function ChkCloseDate(strCloseDate As String, _
    strCompCode As String, strServiceType As String) As Boolean
    On Error GoTo ChkErr
    Dim intDayCut As Integer
    Dim strTranDate As String
    Dim intPara6 As Integer
        intDayCut = Val(GetSystemParaItem("DayCut", strCompCode, "", "SO041") & "")
        strTranDate = GetRsValue("Select tranDate from " & GetOwner & "So062 Where CompCode= " & strCompCode & " And Type=1 Order By TranDate Desc") & ""
        '���O�Ѽ��ɡ]SO043)�A�����O��n������<Para6>�A�ѫ���ϥΡI
        intPara6 = Val(GetSystemParaItem("Para6", strCompCode, strServiceType, "SO043") & "")
    
        If strCloseDate = "" Then If vbNo = MsgBox("�����O�_���ŭȡI", vbExclamation, "�T���I") Then Exit Function
        
        If Not IsDate(strCloseDate) Then
            MsgBox "������X�k�I", vbExclamation, "�T���I"
            Exit Function
        End If
        
        If CDate(strCloseDate) > RightDate Then MsgBox "������W�L���Ѥ���I", vbExclamation, "�T���I": Exit Function
        
        If (RightDate - CDate(strCloseDate)) > intPara6 Then
            MsgBox "������w�W�L�t�γ]�w���w�������I", vbExclamation, "�T���I":  Exit Function
        End If
        If CDate(strCloseDate) <= IIf(strTranDate <> "", CDate(strTranDate), 0) Then
            If intDayCut = 1 And CDate(strCloseDate) = IIf(strTranDate <> "", CDate(strTranDate), 0) Then ChkCloseDate = True: Exit Function
            MsgBox strTranDate & "���e�w���L�鵲�A���i�ϥΤ��e����J�b", vbExclamation, "�T���I"
            Exit Function
        End If
        ChkCloseDate = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "ChkCloseDate"
End Function

Private Sub gilBankCode_Change()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strEmpNo As String
        If Len(gilBankCode.GetCodeNo) > 0 And Len(gilBankCode.GetDescription) > 0 Then
            GetInitData
            strEmpNo = GetRsValue("Select EmpNo From " & GetOwner & "CD018 Where CodeNo =" & gilBankCode.GetCodeNo) & ""
            If strEmpNo <> "" Then
                gilClctEn.SetCodeNo strEmpNo
                gilClctEn.Query_Description
            End If
        End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gilBankCode_Change")
End Sub

Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilBankCode, "CodeNo", "Description", "CD018", , , , , , , " Where UPPER(PrgName) like '%ACHTRANREFER%'", True
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)
        SetgiList gilCMCode, "CodeNo", "Description", "CD031", , , , , 3, 20, , True
        SetgiList gilClctEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
        
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGil"
End Sub

Private Sub gilCompCode_Change()
    On Error GoTo ChkErr
    garyGi(16) = gilCompCode.GetOwner
    Dim strErrFile As String
        If Len(gilCompCode.GetCodeNo & "") = 0 Then gilCompCode.SetFocus: strErrFile = "���q�O": GoTo 99
        GiListFilter gilBankCode, , gilCompCode.GetCodeNo
        gilBankCode.Filter = gilBankCode.Filter & IIf(gilBankCode.Filter = "", " Where ", " And ") & " UPPER(PrgName) like '%ACHTRANREFER%'"
        Call subGil
    Exit Sub
99:
  MsgMustBe (strErrFile)
  Exit Sub
ChkErr:
  ErrSub Me.Name, "gilCompCode_Change"
End Sub

Private Function GetLogFileName(strPath As String) As String
    On Error Resume Next
        GetLogFileName = strPath & IIf(Right(strPath, 1) = "\", "", "\") & "ACHTranReferIn.log"
End Function

Public Function GetOwner(Optional strOwner As String) As String
  On Error Resume Next
    If strOwner = "" Then strOwner = garyGi(16)
    If strOwner <> "" Then GetOwner = strOwner & "."
End Function

Private Sub subGim()
    On Error Resume Next
        '���O����
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "�N�X", "�W��", , True
End Sub

