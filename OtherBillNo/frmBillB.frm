VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBillB 
   BorderStyle     =   1  '單線固定
   Caption         =   "匯入資料 [frmBillB]"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5115
   Icon            =   "frmBillB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5115
   StartUpPosition =   1  '所屬視窗中央
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   90
      TabIndex        =   6
      Top             =   1890
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.匯入"
      Height          =   405
      Left            =   90
      TabIndex        =   5
      Top             =   1320
      Width           =   1005
   End
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
      Left            =   1290
      TabIndex        =   4
      Top             =   1320
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   4935
      Begin VB.TextBox txtFile 
         Height          =   345
         Left            =   1050
         TabIndex        =   2
         Top             =   420
         Width           =   2955
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "...."
         Height          =   375
         Left            =   4080
         TabIndex        =   1
         Top             =   420
         Width           =   555
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
         TabIndex        =   3
         Top             =   480
         Width           =   1065
      End
   End
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   4200
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBillB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const FXlsFields As String = " * "
Private rsXls As New ADODB.Recordset
Private strErrLog As String
Private objStream As New FileSystemObject
Private Const strLogName As String = "匯入異常記錄.Txt"
Private lngRecCount As Long
Private lngImportCount As Long
Private lngErrCount As Long
Private rs315 As New ADODB.Recordset
Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub cmdOK_Click()
 On Error GoTo ChkErr
    If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    If Not GetXlsData(txtFile.Text) Then
        MsgBox "匯入失敗！", vbInformation, "訊息"
        Screen.MousePointer = vbDefault
        Unload Me
    Else
        If rsXls.EOF Then
            MsgBox "無任何資料！", vbInformation, "訊息"
            Screen.MousePointer = vbDefault
            Exit Sub
            
        End If
        If ImportData Then
            MsgBox "匯入完成！" & vbCrLf & " 總筆數 : " & lngRecCount & vbCrLf & _
                    " 匯入筆數 : " & lngImportCount & vbCrLf & _
                    " 錯誤筆數 : " & lngErrCount, vbInformation, "訊息"
            If lngErrCount > 0 Then
                Shell "NOTEPAD " & strErrLog & "\" & strLogName, vbNormalFocus
            End If
        
            Unload Me
        Else
            MsgBox "匯入失敗！", vbCritical, "訊息"
            If lngErrCount > 0 Then
                Shell "NOTEPAD " & strErrLog & "\" & strLogName, vbNormalFocus
            End If
        End If
    End If
    ProgressBar1.Value = 0
    Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
  ErrSub Me.Name, "cmdOK_Click"
End Sub

Private Sub cmdOpen_Click()
On Error GoTo ChkErr
    With comdPath
        .DialogTitle = "選擇輸出路徑"
        .Filter = "Microsoft Excel 活頁簿 (*.XLS)|*.XLS"
        .Action = 1
         If .FileName <> "" Then txtFile = .FileName
    End With
    Exit Sub
ChkErr:
   ErrSub Me.Name, "cmdOpen_Click"
End Sub
Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
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
ChkErr:
  ErrSub Me.Name, "IsDataOK"
End Function
Private Function GetXlsData(ByVal aFile As String) As Boolean
  On Error GoTo ChkErr
    
    Dim strTmp As String
    
    Dim i As Long
    Set rsXls = GetXlsRs(aFile, "Sheet1", FXlsFields)
    If rsXls Is Nothing Then Exit Function
    If rsXls.State = adStateClosed Then Exit Function
    'If rsXls.EOF Then Exit Function
    GetXlsData = True
    Exit Function
ChkErr:
  ErrSub Me.Name, "GetXlsData"
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
    If KeyCode = vbKeyF2 Then
        cmdOK.Value = True
    End If
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    strErrLog = App.Path
    If Right(strErrLog, 1) <> "\" Then
        strErrLog = strErrLog & "\"
    End If
    strErrLog = strErrLog & "ErrLog"
    If Not objStream.FolderExists(strErrLog) Then
        objStream.CreateFolder (strErrLog)
    End If
    strErrLog = strErrLog & "\" & Format(Now, "yyyymmdd")
    If Not objStream.FolderExists(strErrLog) Then
        objStream.CreateFolder (strErrLog)
    End If
    
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    Set frmBillA.uRs = rs315
    CloseRecordset rsXls
    Set objStream = Nothing
End Sub
Private Function ImportData() As Boolean
  On Error GoTo ChkErr
    Dim aSO315SQL As String
    
    Dim aMsg As String
    Dim blnTrans As Boolean
    Dim i As Long
    aSO315SQL = "SELECT * FROM " & "SO315 WHERE 1=0"
    lngErrCount = 0
    lngImportCount = 0
    lngRecCount = 0
    If Not GetRS(rs315, aSO315SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsXls.RecordCount > 0 Then rsXls.MoveFirst
    lngRecCount = rsXls.RecordCount
    ProgressBar1.Min = 0
    ProgressBar1.Max = lngRecCount / 100
    
    Do While Not rsXls.EOF
        aMsg = Empty
        aMsg = chkXLSDataOK
        ProgressBar1.Value = rsXls.AbsolutePosition / 100
        blnTrans = False
        If aMsg = "" Then
            If Not IsXLSDouble Then
                On Error GoTo lNext
                gcnGi.BeginTrans
                blnTrans = True
                With rs315
                    .AddNew
                    For i = 0 To rsXls.Fields.Count - 1
                        If Len(rsXls(i) & "") > 0 Then
                            .Fields(rsXls.Fields(i).Name) = rsXls.Fields(i) & ""
                        End If
                        
                    Next
                    .Update
                    lngImportCount = lngImportCount + 1
                End With
                gcnGi.CommitTrans
            Else
                aMsg = "第 " & rsXls.AbsolutePosition & " 筆資料 [廠商別 : " & rsXls("VenderCode") & "]" & _
                    "、[帳單號碼 : " & rsXls("VBillNo") & "] 資料重複！無法匯入"
                Call WriteErrLog(aMsg)
                lngErrCount = lngErrCount + 1
                If blnTrans Then gcnGi.RollbackTrans
            End If
        Else
            Call WriteErrLog(aMsg)
            lngErrCount = lngErrCount + 1
            If blnTrans Then gcnGi.RollbackTrans
        End If
lNext:
        If Err.Number <> 0 Then
            rs315.CancelUpdate
            If blnTrans Then gcnGi.RollbackTrans
            aMsg = "第 " & rsXls.AbsolutePosition & " 筆資料 第 " & i + 1 & " 欄位有問題！無法匯入"
            Call WriteErrLog(aMsg)
            lngErrCount = lngErrCount + 1
            
            Err.Clear
        End If
        
        rsXls.MoveNext
    Loop
    
    ImportData = True
    Exit Function
ChkErr:
  
  Call ErrSub(Me.Name, "ImportData")
End Function
Private Sub WriteErrLog(ByVal aMsgTxt As String)
  On Error GoTo ChkErr
    Dim aFile As String
    Dim aMsg As String
    aFile = strErrLog & "\" & strLogName
    aMsg = "---------------------------------------------------------------------------------" & vbCrLf & _
        "錯誤時間: " & Now & vbCrLf & _
        "錯誤原因: " & aMsgTxt & vbCrLf & _
        "---------------------------------------------------------------------------------"
    Dim txtStream As TextStream
    
    Set txtStream = objStream.OpenTextFile(aFile, ForAppending, True)
    txtStream.WriteLine aMsg
    On Error Resume Next
    
    txtStream.Close
    Set txtStream = Nothing
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "WriteErrLog")
End Sub
Private Function IsXLSDouble() As Boolean
    On Error GoTo ChkErr
    Dim rschk As New ADODB.Recordset
    Dim aSQL As String
    aSQL = "SELECT COUNT(1) FROM SO315 WHERE VENDERCODE = '" & rsXls("VENDERCODE") & "'" & _
            " AND VBILLNO ='" & rsXls("VBILLNO") & "'"
    If gcnGi.Execute(aSQL)(0) > 0 Then
        
        IsXLSDouble = True
    
    End If
    On Error Resume Next
    Call CloseRecordset(rschk)
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsXLSDouble")
End Function
Private Function chkXLSDataOK() As String
  On Error GoTo ChkErr
    Dim aRet As String
    aRet = Empty
    If Len(rsXls("VENDERCODE") & "") = 0 Then
        aRet = "第 " & rsXls.AbsolutePosition & " 筆資料 無 [廠商別] 資料！無法匯入"
    End If
    If Len(rsXls("VBILLNO") & "") = 0 Then
        aRet = "第 " & rsXls.AbsolutePosition & " 筆資料 無 [帳單號碼] 資料！無法匯入"
    End If
    If Len(rsXls("CITEMNAME") & "") = 0 Then
        aRet = "第 " & rsXls.AbsolutePosition & " 筆資料 無 [繳費項目] 資料！無法匯入"
    End If
    If Len(rsXls("SHOULDAMT") & "") = 0 Then
        aRet = "第 " & rsXls.AbsolutePosition & " 筆資料 無 [繳費金額] 資料！無法匯入"
    End If
    If Len(rsXls("COMMISSION") & "") = 0 Then
        aRet = "第 " & rsXls.AbsolutePosition & " 筆資料 無 [手續費] 資料！無法匯入"
    End If
    If Len(rsXls("CUSTNAME") & "") = 0 Then
        aRet = "第 " & rsXls.AbsolutePosition & " 筆資料 無 [客戶姓名] 資料！無法匯入"
    End If
    If Len(rsXls("Phone") & "") = 0 Then
        aRet = "第 " & rsXls.AbsolutePosition & " 筆資料 無 [電話] 資料！無法匯入"
    End If
    If gcnGi.Execute("SELECT COUNT(1) FROM " & "SO315A WHERE VenderCode = '" & rsXls("VenderCode") & "'")(0) = 0 Then
        aRet = "第 " & rsXls.AbsolutePosition & " 筆資料 廠商別尚未登入資料！無法匯入"
    End If
    chkXLSDataOK = aRet
    Exit Function
ChkErr:
Call ErrSub(Me.Name, "chkXLSDataOK")
End Function
