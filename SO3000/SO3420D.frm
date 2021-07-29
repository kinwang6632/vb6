VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSO3420D 
   BorderStyle     =   1  '單線固定
   Caption         =   "匯入外部資料[SO3420D]"
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
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
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
      Caption         =   "F2.匯入"
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
         Caption         =   "資料位置"
         BeginProperty Font 
            Name            =   "新細明體"
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
'#5324 增加外部匯入資料 By Kin 2010/03/19
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
                MsgBox "匯入失敗！", vbInformation, "訊息"
                Screen.MousePointer = vbDefault
                Unload Me
            Else
                frmSO3420A.uBillNos = strBillNos
                MsgBox "匯入完成！", vbInformation, "訊息"
                Screen.MousePointer = vbDefault
                Unload Me
            End If
        Case 1
              If Not GetXlsData2(txtFile.Text) Then
                MsgBox "匯入失敗！", vbInformation, "訊息"
                Screen.MousePointer = vbDefault
                Unload Me
            Else
                frmSO3420A.uBillNos = strBillNos
                Set frmSO3420A.uImportExcleData = importExcelData
                MsgBox "匯入完成！", vbInformation, "訊息"
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
        .DialogTitle = "選擇輸出路徑"
        .Filter = "Microsoft Excel 活頁簿 (*.XLS)|*.XLS"
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
        MsgBox "請輸入檔案名稱！", vbInformation, "訊息"
        Exit Function
    End If
    If Dir(txtFile.Text) = "" Then
        MsgBox "檔案不存在！", vbInformation, "訊息"
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
            Case "單據編號"
                iAllFind = iAllFind + 1
            Case "拆機原因"
                iAllFind = iAllFind + 1
            Case "客戶流向"
                iAllFind = iAllFind + 1
            Case "未收原因"
                iAllFind = iAllFind + 1
            Case Else
                
        End Select
    Next
    If iAllFind <> 4 Then
        MsgBox "匯入欄位有誤！", vbCritical, "訊息"
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
            If Len(rsExcel.Fields("單據編號") & "") = 0 Then
                aErrMsg = aErrMsg & "第 " & i + 1 & " 筆 [單據編號] 無值!" & vbCrLf
            End If
            If Len(rsExcel.Fields("拆機原因") & "") = 0 Then
                aErrMsg = aErrMsg & "第 " & i + 1 & " 筆 [拆機原因] 無值!" & vbCrLf
            End If
            '#7329 Cancel to check the field by Kin 2016/10/28
'            If Len(rsExcel.Fields("客戶流向") & "") = 0 Then
'                aErrMsg = aErrMsg & "第 " & i + 1 & " 筆 [客戶流向] 無值!" & vbCrLf
'            End If
            If Len(rsExcel.Fields("未收原因") & "") = 0 Then
                aErrMsg = aErrMsg & "第 " & i + 1 & " 筆 [未收原因] 無值!" & vbCrLf
            End If
            strReasonName = GetRsValue("Select Description from " & GetOwner & "CD014 Where CodeNo = " & rsExcel.Fields("拆機原因")) & ""
            If Len(strReasonName) = 0 Then
                aErrMsg = aErrMsg & "第 " & i + 1 & " 筆找不到 [拆機原因] 代碼!" & vbCrLf
            End If
             If Len(rsExcel.Fields("客戶流向") & "") > 0 Then
                strCustRunName = GetRsValue("Select Description from " & GetOwner & "CD096 Where CodeNo = " & rsExcel.Fields("客戶流向")) & ""
             End If
             '#7329 Cancel to check the field by Kin 2016/10/28
            If (Len(strCustRunName) = 0) And (Len(rsExcel.Fields("客戶流向") & "") > 0) Then
                aErrMsg = aErrMsg & "第 " & i + 1 & " 筆找不到 [客戶流向] 代碼!" & vbCrLf
            End If
            strUCName = GetRsValue("Select Description from " & GetOwner & "CD013 Where CodeNo = " & rsExcel.Fields("未收原因")) & ""
            If Len(strUCName) = 0 Then
                aErrMsg = aErrMsg & "第 " & i + 1 & " 筆找不到 [未收原因] 代碼!" & vbCrLf
            End If
            i = i + 1
            aBillNo = Trim(rsExcel.Fields("單據編號"))
            strCustRunCode = rsExcel.Fields("客戶流向") & ""
            If (Len(strCustRunName) = 0) And (Len(strCustRunCode) = 0) Then
                strCustRunName = "X"
                strCustRunCode = "X"
            End If
            If Len(aErrMsg) = 0 Then
                aryData = rsExcel.Fields("拆機原因") & "," & strReasonName & "," & strCustRunCode & "," & strCustRunName & "," & rsExcel.Fields("未收原因") & "," & strUCName
                importExcelData.Add aBillNo, aryData
                If Len(strBillNos) = 0 Then
                    strBillNos = "'" & Trim(rsExcel.Fields("單據編號")) & "'"
                Else
                    strBillNos = strBillNos & ",'" & Trim(rsExcel.Fields("單據編號")) & "'"
                End If
            End If
            rsExcel.MoveNext
        Loop
    End If
    If Len(aErrMsg) > 0 Then
        MsgBox aErrMsg, vbCritical, "訊息"
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
