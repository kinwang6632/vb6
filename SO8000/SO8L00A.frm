VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSO8L00A 
   BorderStyle     =   1  '單線固定
   Caption         =   "語音外撥資料匯入[SO8L00A]"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   Icon            =   "SO8L00A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   435
      Left            =   2010
      TabIndex        =   8
      Top             =   3060
      Width           =   1125
   End
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   4260
      Top             =   3090
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.匯入"
      Enabled         =   0   'False
      Height          =   435
      Left            =   210
      TabIndex        =   7
      Top             =   3060
      Width           =   1125
   End
   Begin VB.Frame fraData 
      Height          =   2835
      Left            =   60
      TabIndex        =   9
      Top             =   120
      Width           =   7005
      Begin VB.CommandButton cmdImportFile 
         Caption         =   "....."
         Height          =   375
         Left            =   6060
         TabIndex        =   6
         Top             =   2130
         Width           =   465
      End
      Begin VB.TextBox txtImportFile 
         Height          =   405
         Left            =   1500
         TabIndex        =   5
         Top             =   2130
         Width           =   4485
      End
      Begin prjGiList.GiList gilServiceCode 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   570
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiList.GiList gilServiceContent 
         Height          =   315
         Left            =   1500
         TabIndex        =   2
         Top             =   960
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiList.GiList gilContentCode 
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Top             =   1350
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F3Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilAcceptEn 
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   1710
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F3Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilServiceType 
         Height          =   315
         Left            =   1500
         TabIndex        =   0
         Top             =   180
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F3Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin VB.Label lblImportFile 
         Caption         =   "匯入檔案"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   510
         TabIndex        =   15
         Top             =   2250
         Width           =   1215
      End
      Begin VB.Label lblServiceType 
         Caption         =   "服  務  別"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   540
         TabIndex        =   14
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblAcceptEn 
         Caption         =   "受理人員"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   540
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblContentCode 
         Caption         =   "申告內容細項"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   1410
         Width           =   1215
      End
      Begin VB.Label lblServiceContent 
         Caption         =   "申告內容"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   540
         TabIndex        =   11
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label lblServiceCode 
         Caption         =   "來電分類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   540
         TabIndex        =   10
         Top             =   630
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmSO8L00A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#3957 新增外撥匯入資料 By Kin 2008/06/30
Option Explicit
Private dtNow As Date
Private rsSo006 As New ADODB.Recordset
Private Sub subGil()
  On Error GoTo ChkErr
    SetLst gilServiceCode, "CodeNo", "Description", 3, 12, , , "CD008", , True      '來電分類
    SetLst gilServiceContent, "CodeNo", "Description", 3, 12, , , "CD008C", , True  '申告內容
    SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E"    '申告內容細項
    SetLst gilAcceptEn, "EmpNo", "EmpName", 3, 12, , , "CM003", , True  '受理人員
    SetLst gilServiceType, "CodeNo", "Description", 3, 12, , , "CD046"  '服務別
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub
Private Function ImpXlsData() As Boolean
  On Error GoTo ChkErr
    Dim rsXls As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim blnTrans As Boolean
    Dim strQry As String
    Dim objErrLog As New FileSystemObject
    Dim tsErrLog As TextStream
    Dim strErrFile As String
    Dim strErrPath As String
    Dim lngErrCount As Long
    Dim strProcResultNo As String
    Dim strIniFile As String
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    '#4064 測試不OK,Ini路徑要抓Bin的目錄 By Kin 2008/10/06
    strErrPath = ReadGICMIS1("SOEXEPATH")
    '#4064 處理代碼改用Ini檔控制 By Kin 2008/09/24
    If Right(strErrPath, 1) = "\" Then
        strErrFile = strErrPath & "SO8L00ALog.txt"
        strIniFile = strErrPath & "SO8L00APara.Ini"
    Else
        strErrFile = strErrPath & "\SO8L00ALog.txt"
        strIniFile = strErrPath & "\SO8L00APara.Ini"
    End If
    '********************************************************************************
    '#4064 增加判斷Ini檔,如果不存在,則建立檔案並給予預設值 By Kin 2008/09/24
    If Dir(strIniFile) = Empty Then
        WritePrivateProfileString ByVal "-1", ByVal "CodeNo", "610", ByVal strIniFile
        WritePrivateProfileString ByVal "1", ByVal "CodeNo", "601", ByVal strIniFile
    End If
    '*********************************************************************************
    Set rsXls = GetXlsRs(txtImportFile.Text, "Sheet1", "[Q_ID],[TEL],[STATUS],[Company],[CustID],[DialDateTime],[DateTime],[Answer]")
    If rsXls.State = 0 Then
        Me.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If Not GetRS(rsSo006, "Select * From " & GetOwner & "SO006 Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    gcnGi.BeginTrans
    Set tsErrLog = objErrLog.CreateTextFile(strErrFile, True)
    blnTrans = True
    lngErrCount = 0
    Do While Not rsXls.EOF
        strQry = Empty
        strProcResultNo = Empty
        '#3957 測試不ok,對方的資料有可能下面全是空白 By Kin 2008/07/23
        If rsXls("CustId") & "" = "" Or rsXls("Company") & "" = "" Then GoTo NextLoop
        strQry = "Select A.CustName,A.Tel1,B.NodeNo From " & GetOwner & "SO001 A," & GetOwner & "SO014 B" & _
                    " Where A.CustId=" & rsXls("CustId") & _
                    " And A.CompCode=" & rsXls("Company") & _
                    " And A.InstAddrNo=B.AddrNo" & _
                    " And A.CompCode=B.CompCode"
        If Not GetRS(rsTmp, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        If rsTmp.EOF Then
            tsErrLog.WriteLine ("客編:" & rsXls("CustId") & "  於客戶主檔不存在！")
            lngErrCount = lngErrCount + 1
        Else
            '*********************************************************************************************
            '#4064原本處理代碼寫死,現在改為讀取Ini檔案,如果找不到則自動填入預設值 By Kin 2008/09/24
            strProcResultNo = ReadFROMINI(strIniFile, Val(rsXls("STATUS") & ""), "CodeNo")
            If Val(rsXls("STATUS")) = 1 Then
                If strProcResultNo = Empty Then
                    strProcResultNo = 601
                    WritePrivateProfileString ByVal "1", ByVal "CodeNo", "601", ByVal strIniFile
                End If
            Else
                If strProcResultNo = Empty Then
                    strProcResultNo = "610"
                    WritePrivateProfileString ByVal "-1", ByVal "CodeNo", "610", ByVal strIniFile
                End If
            End If
            '**********************************************************************************************
            If Not InsertSO006(Val(rsXls("Company")), Val(rsXls("CustId")), _
                        rsXls("DateTime") & "", rsXls("DialDateTime") & "", _
                        strProcResultNo, rsTmp("CustName") & "", rsTmp("Tel1") & "", rsTmp("NodeNo") & "", rsXls("TEL") & "", rsXls("Q_ID") & "", rsXls("Answer") & "") Then GoTo ChkErr
        End If
NextLoop:
        rsXls.MoveNext
    Loop
    gcnGi.CommitTrans
    On Error Resume Next
    MsgBox "完成筆數：" & rsSo006.RecordCount & vbCrLf & "錯誤筆數：" & lngErrCount, vbInformation, "訊息"
    If lngErrCount Then Shell "NotePad " & strErrFile, vbNormalFocus
    
    blnTrans = False
    Call Close3Recordset(rsXls)
    Call Close3Recordset(rsTmp)
    tsErrLog.Close
    Set tsErrLog = Nothing
    Set objErrLog = Nothing
    ImpXlsData = True
    Exit Function
ChkErr:
    If blnTrans Then
        gcnGi.RollbackTrans
    End If
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Call ErrSub(Me.Name, "ImpXlsData")
End Function
Private Function InsertSO006(ByVal iCompCode As Integer, ByVal iCustId As Long, _
                        ByVal sAcceptTime As String, ByVal sHandleTime As String, _
                        ByVal sProcResultNo As String, ByVal sCustName As String, _
                        ByVal sTel As String, ByVal sNodeNo As String, _
                        ByVal sQTel As String, ByVal sQID As String, ByVal sAnswer As String) As Boolean
  On Error GoTo ChkErr
    Dim strProcResultName As String
    Dim rsResult As New ADODB.Recordset
    Dim strAnswer As String
    strProcResultName = Empty
    '***************************************************************************************************************
    '#4064 判斷有沒有回傳結果代碼,如果沒有不要取結果名稱 By Kin 2008/09/24
    If sProcResultNo <> Empty Then
        Set rsResult = gcnGi.Execute("Select Description From " & GetOwner & "CD008B Where CodeNo=" & sProcResultNo)
        If Not rsResult.EOF Then strProcResultName = rsResult(0) & ""
    End If
    '***************************************************************************************************************
    rsSo006.AddNew
    rsSo006("COMPCODE") = iCompCode
    rsSo006("CustID") = iCustId
    rsSo006("AcceptTime") = sAcceptTime
    rsSo006("HandleTime") = sHandleTime
    rsSo006("ProcResultNo") = IIf(sProcResultNo = Empty, Null, sProcResultNo)
    rsSo006("ProcResultName") = IIf(strProcResultName = Empty, Null, strProcResultName)
    rsSo006("ServiceType") = gilServiceType.GetCodeNo
    rsSo006("SeqNo") = GetInvoiceNo2("SO006")
    rsSo006("CustName") = sCustName
    rsSo006("Tel1") = IIf(sTel = Empty, Null, sTel)
    rsSo006("NodeNo") = IIf(sNodeNo = Empty, Null, sNodeNo)
    rsSo006("ServiceCode") = gilServiceCode.GetCodeNo
    rsSo006("ServiceName") = gilServiceCode.GetDescription
    rsSo006("ServDescCode") = gilServiceContent.GetCodeNo
    rsSo006("ServDescName") = gilServiceContent.GetDescription
    rsSo006("ContentCode") = IIf(gilContentCode.GetCodeNo & "" = "", Null, gilContentCode.GetCodeNo)
    rsSo006("ContentName") = IIf(gilContentCode.GetDescription & "" = "", Null, gilContentCode.GetDescription)
    rsSo006("AcceptEn") = gilAcceptEn.GetCodeNo
    rsSo006("AcceptName") = gilAcceptEn.GetDescription
    rsSo006("UpdEn") = garyGi(1)
    rsSo006("UpdTime") = GetDTString(CStr(dtNow))
    '#4236 如果Answer空白也要填入SO106.Note By Kin 2008/12/25
    'If sQID <> Empty And sAnswer <> Empty Then
        sAnswer = Replace(sAnswer, ",", StrConv(",", vbWide))
        '#3957 測試不OK,要多加中文說明 By Kin 2008/07/29
        strAnswer = "外撥電話:" & sQTel & vbCrLf & "問卷代號:" & sQID & vbCrLf & "問卷答案:" & sAnswer
        rsSo006("Note") = strAnswer
    'End If
    rsSo006.Update
    InsertSO006 = True
    On Error Resume Next
    Call Close3Recordset(rsResult)
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "InsertSO006")
End Function

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub cmdImportFile_Click()
  On Error GoTo ChkErr
    SelectImportFile
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdImportFile_Click")
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    dtNow = Now
    If Not IsDataOk Then Exit Sub
    If Not ChkDirExist(txtImportFile.Text) Then
        MsgBox "所選擇的來源活頁簿(Excel)路徑不正確或檔案不存在! 請確認!", vbInformation, "訊息"
        Exit Sub
    End If
    If Not ImpXlsData Then MsgBox "匯入資料失敗！", vbCritical, "訊息"
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyEscape Then Unload Me: Exit Sub
    If KeyCode = vbKeyF2 Then
        If cmdOK.Enabled Then cmdOK.Value = True
    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGil
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  Call Close3Recordset(rsSo006)
End Sub
Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    IsDataOk = False
    If gilServiceCode.GetCodeNo & "" = "" Then MsgBox "來電分類必須有值！", vbExclamation, "訊息": gilServiceCode.SetFocus: Exit Function
    If gilServiceContent.GetCodeNo & "" = "" Then MsgBox "申告內容必須有值！", vbExclamation, "訊息": gilServiceContent.SetFocus: Exit Function
    If gilAcceptEn.GetCodeNo & "" = "" Then MsgBox "受理人員必須有值！", vbExclamation, "訊息": gilAcceptEn.SetFocus: Exit Function
    If gilServiceType.GetCodeNo & "" = "" Then MsgBox "服務類別必須有值！", vbExclamation, "訊息": gilServiceType.SetFocus: Exit Function
    IsDataOk = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDateOK")
End Function
Private Sub gilServiceCode_Change()
  On Error GoTo ChkErr
    
    If Len(Trim(gilServiceCode.GetCodeNo)) > 0 Then

        SetgiList gilServiceContent, "CodeNo", "Description", "CD008A", , , , , , , " Where CodeNo in (Select CodeNo From " & GetOwner & "CD008C " & _
                                                                " Where ServiceCode=" & gilServiceCode.GetCodeNo & _
                                                                " )" & IIf(gilServiceType.GetCodeNo <> "", " And (ServiceType='" & gilServiceType.GetCodeNo & "' or ServiceType is Null)", ""), True
    Else
        gilServiceContent.Filter = ""
        SetgiList gilServiceContent, "CodeNo", "Description", "CD008A", , , , , , , " Where 1=0, True"
    
    End If
                                                                    

    If gilServiceContent.GetCodeNo <> "" And gilServiceCode.GetCodeNo <> "" Then
        gilContentCode.Filter = ""

        SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", " WHERE " & _
                                                                   "Exists (Select * From CD008D Where CD008E.CodeNo=CD008D.CodeNo And " & _
                                                                               "CD008E.ServiceType=CD008D.ServiceType And CD008D.StopFlag<>1 or CD008E.ServiceType is Null)" & _
                                                                                IIf(gilServiceType.GetCodeNo <> "", " And (ServiceType='" & gilServiceType.GetCodeNo & "' or ServiceType is Null)", "")
'        SetLst gilContent, "CodeNo", "Description", 3, 12, , , "CD008E", _
'                " WHERE " & Filter8E8D(strSvcType) & " AND NVL(SERVICETYPE,'" & strSvcType & "' ) ='" & strSvcType & "'", False
    Else
        gilContentCode.Filter = ""
        SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", "Where 1=0"
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gilServiceCode_Change"
End Sub

Private Sub SelectImportFile()
  On Error GoTo ChkErr
    With comdPath
        .Filter = "Excel檔案|*.XLS"
        .Action = 1
        If .FileName <> "" Then txtImportFile.Text = .FileName
    End With
    
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "SelectImportFile")
End Sub
Private Sub gilServiceContent_Change()
  On Error GoTo ChkErr
    If gilServiceContent.GetCodeNo <> "" Then
        gilContentCode.Filter = ""

'        SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", " WHERE " & _
'                                                                   "Exists (Select * From CD008D Where CD008E.CodeNo=CD008D.CodeNo And " & _
'                                                                               "CD008E.ServiceType=CD008D.ServiceType And CD008D.StopFlag<>1 or CD008E.ServiceType is Null)" & _
'                                                                               IIf(gilServiceType.GetCodeNo <> "", " And (ServiceType='" & gilServiceType.GetCodeNo & "' or ServiceType is Null)", "")
        SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", _
                " WHERE " & Filter8E8D(gilServiceType.GetCodeNo) & " AND NVL(SERVICETYPE,'" & gilServiceType.GetCodeNo & "' ) ='" & gilServiceType.GetCodeNo & "'", False
                                                   
    Else
        gilContentCode.Filter = ""
        SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", "Where 1=0"
    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceContent_Change")
End Sub
Private Function Filter8E8D(strSvcType As String) As String
  On Error Resume Next
    Dim s As String
    s = "SELECT B.CODENO FROM " & GetOwner & "CD008E A," & _
                                                        GetOwner & "CD008D B" & _
                                                        " WHERE A.CONTENTCODE=" & gilServiceContent.GetCodeNo & _
                                                        IIf(strSvcType <> Empty, " AND NVL(B.SERVICETYPE,'" & strSvcType & "')='" & strSvcType & "'", "") & _
                                                        " AND B.CODENO=A.CODENO" & _
                                                        " AND B.STOPFLAG <>1"
    Filter8E8D = RPxx(gcnGi.Execute("SELECT B.CODENO FROM " & GetOwner & "CD008E A," & _
                                                        GetOwner & "CD008D B" & _
                                                        " WHERE A.CONTENTCODE=" & gilServiceContent.GetCodeNo & _
                                                        IIf(strSvcType <> Empty, " AND NVL(B.SERVICETYPE,'" & strSvcType & "')='" & strSvcType & "'", "") & _
                                                        " AND B.CODENO=A.CODENO" & _
                                                        " AND B.STOPFLAG <>1").GetString(adClipString, , , ",", ""))
    If Err.Number = 3021 Then
        Err.Clear
        Filter8E8D = " 0=1"
        'gilContentCode.Enabled = False
        Exit Function
    ElseIf Err.Number <> 0 Then
        ErrSub Name, "Filter8E8D"
    End If
    
    If Right(Filter8E8D, 1) = "," Then Filter8E8D = Left(Filter8E8D, Len(Filter8E8D) - 1)
    If Filter8E8D = Empty Then
        Filter8E8D = " 0=1"
        'gilContentCode.Enabled = False
    Else
        Filter8E8D = " CODENO IN (" & Filter8E8D & " ) AND CONTENTCODE=" & gilServiceContent.GetCodeNo
        'gilContentCode.Enabled = True
    End If
End Function
Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    SetLst gilServiceCode, "CodeNo", "Description", 3, 12, , , "CD008", _
        IIf(gilServiceType.GetCodeNo <> "", " Where (ServiceType='" & gilServiceType.GetCodeNo & "' Or ServiceType is Null) And StopFlag<>1 ", " Where StopFlag<>1")
    Call gilServiceCode_Change
    
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

Private Sub txtImportFile_Change()
  On Error GoTo ChkErr:
    If txtImportFile.Text <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "txtImportFile_Change")
End Sub
