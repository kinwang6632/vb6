VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.0#0"; "GiTime.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Begin VB.Form frmSO1625A 
   Caption         =   "CA命令資料管理 [SO1625A]"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO1625A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7650
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   12000
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdFind 
      Caption         =   "F3. 尋找"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "F4. 清除"
      Height          =   375
      Left            =   1245
      TabIndex        =   5
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   345
      Left            =   10500
      TabIndex        =   7
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "預約完成命令轉入歷史檔"
      Height          =   375
      Left            =   4350
      TabIndex        =   6
      Top             =   7200
      Width           =   2475
   End
   Begin VB.Frame fraData 
      Height          =   1395
      Left            =   120
      TabIndex        =   9
      Top             =   450
      Width           =   11655
      Begin Gi_Multi.GiMulti gimStatus 
         Height          =   345
         Left            =   5730
         TabIndex        =   12
         Top             =   300
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   609
         ButtonCaption   =   "命令狀態"
      End
      Begin Gi_Time.GiTime gdtProcessingDate1 
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         Top             =   330
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Time.GiTime gdtProcessingDate2 
         Height          =   315
         Left            =   3570
         TabIndex        =   2
         Top             =   330
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Multi.GiMulti gimCommand 
         Height          =   345
         Left            =   360
         TabIndex        =   13
         Top             =   810
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   609
         ButtonCaption   =   "預約命令"
      End
      Begin VB.Label lblProcessingDate 
         AutoSize        =   -1  'True
         Caption         =   "預約時間"
         Height          =   195
         Left            =   420
         TabIndex        =   11
         Top             =   390
         Width           =   780
      End
      Begin VB.Label lblTo1 
         AutoSize        =   -1  'True
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3330
         TabIndex        =   10
         Top             =   390
         Width           =   195
      End
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   5235
      Left            =   90
      TabIndex        =   3
      Top             =   1920
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   9234
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   285
      Left            =   1065
      TabIndex        =   0
      Top             =   90
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   750
      FldWidth2       =   1600
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司名稱"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   165
      TabIndex        =   8
      Top             =   150
      Width           =   780
   End
End
Attribute VB_Name = "frmSO1625A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const TableName = "SOAC0202"
Private Const SendTableName = "Send_Nagra"
Dim rs As New ADODB.Recordset

Private Sub cmdClear_Click()
    Call OpenData(True)

End Sub

Private Sub cmdExecute_Click()
    On Error GoTo ChkErr
    Dim strModeType As String
    Dim varX() As String, intLoop As Integer
    Dim strChCode As String
    Dim strField As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngCount As Long
        With rs
            If .RecordCount = 0 Then Exit Sub
            .MoveFirst
            On Error GoTo Err
            gcnGi.BeginTrans
            Do While Not .EOF
                If .Fields("CMD_Status") = "C" And .Fields("ProcessingDate") & "" <> "" Then
                    If Not GetRS(rsTmp, "Select * From " & GetOwner & TableName & "tmp Where SeqNo =" & .Fields("SeqNo")) Then GoTo Err
                    If Not rsTmp.EOF Then
                        '先更新SOAC0x02
                        strField = "AuthorStartDate, AuthorStopDate, ChCode, ChName, CompCode, CustId, IPPVCredit, ModeType, SmartCardNo, STBSNo, UpdEn, UpdTime, ResvTime, BillNo"
                        If Not ExecuteCommand("Insert Into " & GetOwner & TableName & " (" & strField & ") Select " & strField & " From " & GetOwner & TableName & "tmp Where SeqNo = " & .Fields("SeqNo")) Then GoTo Err
                        '刪除相關資料
                        If Not ExecuteCommand("Delete From " & GetOwner & TableName & "tmp Where SeqNo = " & .Fields("SeqNo")) Then GoTo Err
                        '更新SO005
                        Select Case .Fields("HIGH_LEVEL_CMD_ID") & ""
                            '開頻道,暫停頻道,恢復頻道,期限延展,收視權限複製
                            Case "B1", "B4", "B5", "B6", "B7"
                                If Not InsertToSO005(.Fields("SeqNo")) Then GoTo Err
                            '開機,關機,關頻道(全部)
                            Case "A1", "A2", "B3"
                                '刪除頻道
                                If Not ExecuteCommand("Delete From " & GetOwner & "SO005 Where CVTSNo = '" & rsTmp("STBSNO") & "' And CompCode = " & rsTmp("CompCode")) Then GoTo Err
                                '回填認證電話
                                If .Fields("HIGH_LEVEL_CMD_ID") = "A1" Then
                                    If Not ExecuteCommand("Update " & GetOwner & "SO004 Set RecallTel='" & rs("PHONE_NUM_1") & "' Where CustId = " & rsTmp("CustId") & " And FaciSNo = '" & rsTmp("STBSNO") & "' And CompCode = " & rsTmp("CompCode")) Then GoTo Err
                                End If
                            '關頻道(個別)
                            Case "B2"
                                varX = Split(.Fields("Notes") & "", ",")
                                For intLoop = 0 To UBound(varX)
                                    strChCode = GetChCode(varX(intLoop))
                                    If Not ExecuteCommand("Delete From " & GetOwner & "So005 Where CvtSNo = '" & rsTmp("STBSNO") & "' And ChCode = '" & strChCode & "'") Then GoTo Err
                                Next
                            '開試看頻道
                            Case "BA"
                            '關試看頻道
                            Case "BB"
                            '設定STB認證電話(回撥用)
                            Case "C3"
                                If Not ExecuteCommand("Update " & GetOwner & "SO004 Set RecallTel='" & rs("PHONE_NUM_1") & "' Where CustId = " & rsTmp("CustId") & " And STBSNo = '" & rsTmp("STBSNO") & "' And CompCode = " & rsTmp("CompCode")) Then GoTo Err
                            '設定親子密碼
                            Case "E1"
                            '設定定期驗證子機
                            Case "E5"
                            '取消子母機驗證
                            Case "E6"
                            '臨時驗證子機
                            Case "E7"
                        End Select
                    End If
                    If Not ExecuteCommand("Delete From " & gLogInID & SendTableName & " Where RowId = '" & .Fields("RowId") & "'") Then GoTo Err
                    lngCount = lngCount + 1
                End If
                .MoveNext
            Loop
            gcnGi.CommitTrans
            MsgBox "轉入完成,共轉入" & lngCount & "筆!!", vbInformation, gimsgPrompt
        End With
        Call OpenData(True)
    Exit Sub
Err:
    gcnGi.RollbackTrans
    MsgBox "更新失敗!!請查證!!", vbCritical, gimsgWarning
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdExecute_Click"
End Sub

Private Function GetChCode(strChannelId As String) As String
    On Error Resume Next
        GetChCode = GetRsValue("Select CodeNo From " & GetOwner & "CD024 Where ChannelId = '" & strChannelId & "'") & ""
End Function

Private Function InsertToSO005(lngSeqNo As Long) As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim intLoop As Integer, strChCode As String
    Dim strField As String
        If Not GetRS(rsTmp, "Select * From " & GetOwner & "SO005tmp Where SeqNo = " & lngSeqNo) Then Exit Function
        Do While Not rsTmp.EOF
            strChCode = strChCode & ",'" & rsTmp("ChCode") & "'"
            strField = "ChCode, CitemCode, CompCode, CustId, CvtId, CvtSNo, DueDate, OrderNo, ServiceType, SetEn, SetName, SetTime, SmartCardNo, Status"
            If Not ExecuteCommand("Delete From " & GetOwner & "SO005 Where CustId = " & rsTmp("CustId") & " And CvtSNo = '" & rsTmp("CvtSNO") & "' And ChCode = '" & rsTmp("ChCode") & "'") Then Exit Function
            If Not InsertToOracle("SO005", rsTmp) Then Exit Function
            'If Not ExecuteCommand("Insert Into " & GetOwner & "SO005 (" & strField & ") Select " & strField & " From " & GetOwner & "SO005tmp Where SeqNo = " & lngSeqNo & " And ChCode = '" & rsTmp("ChCode") & "')") Then Exit Function
            rsTmp.MoveNext
        Loop
        strChCode = Mid(strChCode, 2)
        If strChCode <> "" Then
            If Not ExecuteCommand("Delete From " & GetOwner & "SO005tmp Where SeqNo = " & lngSeqNo) Then Exit Function
        End If
        Call CloseRecordset(rsTmp)
        InsertToSO005 = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "InsertToSO005"
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Call OpenData(False)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        Select Case KeyCode
            Case vbKeyEscape
                Unload Me
            Case vbKeyF3
                If cmdFind.Enabled Then cmdFind.Value = True
            Case vbKeyF4
                If cmdClear.Enabled Then cmdClear.Value = True
        End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
        Call InitData
        Call subGil
        Call subGim
        Call subGrd
        Call OpenData(True)
        Set objActForm = Me
End Sub

Private Sub InitData()
    On Error Resume Next
        gLogInID = GetSystemParaItem("LogInID", gilCompCode.GetCodeNo, "", "SO041", , True) & ""
        If gLogInID <> "" Then gLogInID = gLogInID & "."

End Sub

Public Sub FindGo()
    Call OpenData(False)
End Sub

Public Sub CancelGo()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        ReleaseCOM Me

End Sub

Private Sub gdtProcessingDate2_GotFocus()
    On Error Resume Next
        If gdtProcessingDate2.GetValue = "" Then gdtProcessingDate2.SetValue RightNow
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
    If giFld.FieldName = "HIGH_LEVEL_CMD_ID" Then
        Value = GetModeTypeValue(Value)
    End If
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        garyGi(16) = gilCompCode.GetOwner
End Sub

Private Function IsDataOk(ByRef strWhere As String) As Boolean
    On Error GoTo ChkErr
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        'If Not MustExist(gdtProcessingDate1, 1, "預約時間起始") Then Exit Function
        'If Not MustExist(gdtProcessingDate2, 1, "預約時間截止") Then Exit Function
        If gilCompCode.GetCodeNo <> "" Then strWhere = strWhere & " And CompCode = " & gilCompCode.GetCodeNo
        If gdtProcessingDate1.GetValue <> "" Then strWhere = strWhere & " And ProcessingDate >=" & GetNullString(gdtProcessingDate1.GetValue(True), giDateV, , True)
        If gdtProcessingDate2.GetValue <> "" Then strWhere = strWhere & " And ProcessingDate <=" & GetNullString(gdtProcessingDate2.GetValue(True), giDateV, , True)
        If gimCommand.GetQryStr <> "" Then strWhere = strWhere & " And HIGH_LEVEL_CMD_ID " & gimCommand.GetQryStr
        If gimStatus.GetQryStr <> "" Then strWhere = strWhere & " And Cmd_Status " & gimStatus.GetQryStr
        If strWhere <> "" Then strWhere = " Where " & Mid(strWhere, 5)
        IsDataOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Private Sub OpenData(blnClear As Boolean)
    On Error GoTo ChkErr
    Dim strWhere As String
    Dim strField As String
        If blnClear Then
            strWhere = " Where 1 = 0 "
        Else
            If Not IsDataOk(strWhere) Then Exit Sub
        End If
        strField = "A.RowId,A.*"
        If Not GetRS(rs, "Select " & strField & " From " & gLogInID & SendTableName & " A " & strWhere & " Order By SeqNo") Then Exit Sub
        Set ggrData.Recordset = rs
        ggrData.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 20, GetCompCodeFilter(gilCompCode, garyGi(15))
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGim()
    On Error Resume Next
        SetgiMultiAddItem gimStatus, "W,P,C,E", "等待,處理中,完成,錯誤", "代碼", "名稱"
        gimStatus.SelectAll
        SetgiMultiAddItem gimCommand, "A1,A2,A3,A4,A5,B1,B2,B3,B4,B5,B6,B7,B9,C2,C3,E1,E2,E3,E4,E5,E6,E7", _
            "開機,關機,停機,復機,註銷,開頻道,關頻道(個別),關頻道(全部),暫停頻道,恢復頻道," & _
            "期限延展,收視權限複製,重設,清除EMM,設定STB認證電話(回撥用),設定親子密碼,傳送訊息," & _
            "ForceTune,設定網路編號,設定定期驗證子機,取消子母機驗證,臨時驗證子機", "代碼", "名稱"
End Sub

Private Sub subGrd()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "CompCode", , , , , "公司"
        .Add "ICC_NO", , , , , "ICC 序號     "
        .Add "HIGH_LEVEL_CMD_ID", , , , , "高階命令    "
        .Add "SUBSCRIPTION_BEGIN_DATE", giControlTypeDate, , , , "收視起始"
        .Add "SUBSCRIPTION_END_DATE", giControlTypeDate, , , , "收視截止"
        .Add "NOTES", , , , , "備註" & Space(40)
        .Add "CMD_Status", , , , , "狀態"
        .Add "ProcessingDate", giControlTypeTime, , , , "預約時間" & Space(10)
        .Add "OPERATOR", , , , , "操作者"
        .Add "UPDTIME", giControlTypeTime, , , , "異動時間" & Space(10)
    End With
    ggrData.AllFields = mFlds
    ggrData.SetHead
  Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd"
End Sub
