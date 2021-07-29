VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSO3420D 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�פJ�~�����[SO3420D]"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5130
   Icon            =   "SO3420D.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5130
   StartUpPosition =   1  '���ݵ�������
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
      Height          =   405
      Left            =   1440
      TabIndex        =   5
      Top             =   1470
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�פJ"
      Height          =   405
      Left            =   210
      TabIndex        =   4
      Top             =   1470
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   4965
      Begin VB.CommandButton cmdOpen 
         Caption         =   "...."
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   390
         Width           =   555
      End
      Begin VB.TextBox txtFile 
         Height          =   345
         Left            =   1050
         TabIndex        =   2
         Top             =   390
         Width           =   2955
      End
      Begin VB.Label Label1 
         Caption         =   "��Ʀ�m"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   1
         Top             =   450
         Width           =   1065
      End
   End
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   3360
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSO3420D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#5324 �W�[�~���פJ��� By Kin 2010/03/19
Option Explicit

Private strBillNos As String
Private excelType As Integer
Private importExcelData As Dictionary
Private Sub cmdCancel_Click()
  On Error Resume Next
    Screen.MousePointer = vbDefault
    Unload Me
    
    
End Sub

Private Sub cmdOK_Click()
  On Error GoTo chkErr
    If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
    Select Case excelType
        Case 0
            If Not GetXlsData(txtFile.Text) Then
                MsgBox "�פJ���ѡI", vbInformation, "�T��"
                Screen.MousePointer = vbDefault
                Unload Me
            Else
                frmSO3420A.uBillNos = strBillNos
                MsgBox "�פJ�����I", vbInformation, "�T��"
                Screen.MousePointer = vbDefault
                Unload Me
            End If
        Case 1
              If Not GetXlsData2(txtFile.Text) Then
                MsgBox "�פJ���ѡI", vbInformation, "�T��"
                Screen.MousePointer = vbDefault
                Unload Me
            Else
                frmSO3420A.uBillNos = strBillNos
                Set frmSO3420A.uImportExcleData = importExcelData
                MsgBox "�פJ�����I", vbInformation, "�T��"
                Screen.MousePointer = vbDefault
                Unload Me
            End If
    End Select
    
    Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
  ErrSub Me.Name, "cmdOK_Click"
End Sub

Private Sub cmdOpen_Click()
  On Error GoTo chkErr
    With comdPath
        .DialogTitle = "��ܿ�X���|"
        .Filter = "Microsoft Excel ����ï (*.XLS)|*.XLS"
        .Action = 1
         If .FileName <> "" Then txtFile = .FileName
    End With
    Exit Sub
chkErr:
   ErrSub Me.Name, "cmdOpen_Click"
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub
Private Function IsDataOk() As Boolean
  On Error GoTo chkErr
    If txtFile.Text = "" Then
        MsgBox "�п�J�ɮצW�١I", vbInformation, "�T��"
        Exit Function
    End If
    If Dir(txtFile.Text) = "" Then
        MsgBox "�ɮפ��s�b�I", vbInformation, "�T��"
        Exit Function
    End If
    IsDataOk = True
    Exit Function
chkErr:
  ErrSub Me.Name, "IsDataOK"
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = vbKeyF2 Then If cmdOk.Enabled Then cmdOk.Value = True
    If KeyCode = vbKeyEscape Then Unload Me
    Exit Sub
End Sub

Private Sub Form_Load()
  On Error Resume Next
    strBillNos = Empty
End Sub
Private Function addToDictionary(ByRef rsExcel As ADODB.Recordset) As Boolean
  On Error GoTo chkErr
    Dim aErrMsg As String
    Dim i As Integer
    Set importExcelData = New Dictionary
    importExcelData.CompareMode = TextCompare
    Dim aryData As String
    Dim strReasonName As String
    Dim strCustRunName As String
    Dim strCustRunCode As String
    Dim strUCName As String
    Dim aBillNo  As String
    strBillNos = ""
    Dim iAllFind As Integer
    iAllFind = 0
    For i = 0 To rsExcel.Fields.Count - 1
        Select Case rsExcel.Fields(i).Name
            Case "��ڽs��"
                iAllFind = iAllFind + 1
            Case "�����]"
                iAllFind = iAllFind + 1
            Case "�Ȥ�y�V"
                iAllFind = iAllFind + 1
            Case "������]"
                iAllFind = iAllFind + 1
            Case Else
                
        End Select
    Next
    If iAllFind <> 4 Then
        MsgBox "�פJ��즳�~�I", vbCritical, "�T��"
        addToDictionary = False
        Exit Function
    End If
    If rsExcel.RecordCount > 0 Then
        rsExcel.MoveFirst
        i = 0
        Do While Not rsExcel.EOF
            aBillNo = ""
            aryData = ""
            strReasonName = ""
            strCustRunName = ""
            strCustRunCode = ""
            strUCName = ""
            If Len(rsExcel.Fields("��ڽs��") & "") = 0 Then
                aErrMsg = aErrMsg & "�� " & i + 1 & " �� [��ڽs��] �L��!" & vbCrLf
            End If
            If Len(rsExcel.Fields("�����]") & "") = 0 Then
                aErrMsg = aErrMsg & "�� " & i + 1 & " �� [�����]] �L��!" & vbCrLf
            End If
            '#7329 Cancel to check the field by Kin 2016/10/28
'            If Len(rsExcel.Fields("�Ȥ�y�V") & "") = 0 Then
'                aErrMsg = aErrMsg & "�� " & i + 1 & " �� [�Ȥ�y�V] �L��!" & vbCrLf
'            End If
            If Len(rsExcel.Fields("������]") & "") = 0 Then
                aErrMsg = aErrMsg & "�� " & i + 1 & " �� [������]] �L��!" & vbCrLf
            End If
            strReasonName = GetRsValue("Select Description from " & GetOwner & "CD014 Where CodeNo = " & rsExcel.Fields("�����]")) & ""
            If Len(strReasonName) = 0 Then
                aErrMsg = aErrMsg & "�� " & i + 1 & " ���䤣�� [�����]] �N�X!" & vbCrLf
            End If
             If Len(rsExcel.Fields("�Ȥ�y�V") & "") > 0 Then
                strCustRunName = GetRsValue("Select Description from " & GetOwner & "CD096 Where CodeNo = " & rsExcel.Fields("�Ȥ�y�V")) & ""
             End If
             '#7329 Cancel to check the field by Kin 2016/10/28
            If (Len(strCustRunName) = 0) And (Len(rsExcel.Fields("�Ȥ�y�V") & "") > 0) Then
                aErrMsg = aErrMsg & "�� " & i + 1 & " ���䤣�� [�Ȥ�y�V] �N�X!" & vbCrLf
            End If
            strUCName = GetRsValue("Select Description from " & GetOwner & "CD013 Where CodeNo = " & rsExcel.Fields("������]")) & ""
            If Len(strUCName) = 0 Then
                aErrMsg = aErrMsg & "�� " & i + 1 & " ���䤣�� [������]] �N�X!" & vbCrLf
            End If
            i = i + 1
            aBillNo = Trim(rsExcel.Fields("��ڽs��"))
            strCustRunCode = rsExcel.Fields("�Ȥ�y�V") & ""
            If (Len(strCustRunName) = 0) And (Len(strCustRunCode) = 0) Then
                strCustRunName = "X"
                strCustRunCode = "X"
            End If
            If Len(aErrMsg) = 0 Then
                aryData = rsExcel.Fields("�����]") & "," & strReasonName & "," & strCustRunCode & "," & strCustRunName & "," & rsExcel.Fields("������]") & "," & strUCName
                importExcelData.Add aBillNo, aryData
                If Len(strBillNos) = 0 Then
                    strBillNos = "'" & Trim(rsExcel.Fields("��ڽs��")) & "'"
                Else
                    strBillNos = strBillNos & ",'" & Trim(rsExcel.Fields("��ڽs��")) & "'"
                End If
            End If
            rsExcel.MoveNext
        Loop
    End If
    If Len(aErrMsg) > 0 Then
        MsgBox aErrMsg, vbCritical, "�T��"
        importExcelData.RemoveAll
        Set importExcelData = Nothing
        addToDictionary = False
    Else
        'Set frmSO3420A.uImportExcleData = importExcelData
        addToDictionary = True
    End If
    Exit Function
chkErr:

  ErrSub Me.Name, "addToDictionary"
End Function
Private Function GetXlsData2(ByVal aFile As String) As Boolean
  On Error GoTo chkErr
    Dim rsXls As New ADODB.Recordset
    Dim strTmp As String
    Dim strAry() As String
    Dim i As Long
    Set rsXls = GetXlsRs(aFile, "Sheet1", " * ")
    If rsXls Is Nothing Then Exit Function
    If rsXls.State = adStateClosed Then Exit Function
    If rsXls.EOF Then Exit Function
    GetXlsData2 = addToDictionary(rsXls)
    
'    strTmp = rsXls.GetString(, , , ",")
'    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
'    strAry = Split(strTmp, ",")
'    For i = LBound(strAry) To UBound(strAry)
'        If strBillNos <> Empty Then
'            strBillNos = strBillNos & ",'" & strAry(i) & "'"
'        Else
'            strBillNos = "'" & strAry(i) & "'"
'        End If
'    Next i
    'GetXlsData2 = True
On Error Resume Next
    CloseRecordset rsXls
    Set rsXls = Nothing
    Exit Function
chkErr:
  ErrSub Me.Name, "GetXlsData2"
End Function

Private Function GetXlsData(ByVal aFile As String) As Boolean
  On Error GoTo chkErr
    Dim rsXls As New ADODB.Recordset
    Dim strTmp As String
    Dim strAry() As String
    Dim i As Long
    Set rsXls = GetXlsRs(aFile, "Sheet1", " * ")
    If rsXls Is Nothing Then Exit Function
    If rsXls.State = adStateClosed Then Exit Function
    If rsXls.EOF Then Exit Function
    strTmp = rsXls.GetString(, , , ",")
    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    strAry = Split(strTmp, ",")
    For i = LBound(strAry) To UBound(strAry)
        If strBillNos <> Empty Then
            strBillNos = strBillNos & ",'" & strAry(i) & "'"
        Else
            strBillNos = "'" & strAry(i) & "'"
        End If
    Next i
    GetXlsData = True
On Error Resume Next
    CloseRecordset rsXls
    Set rsXls = Nothing
    Exit Function
chkErr:
  ErrSub Me.Name, "GetXlsData"
End Function
Public Property Let uExcelType(ByVal vData As Integer)
    excelType = vData
End Property
