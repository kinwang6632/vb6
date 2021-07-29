VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExeCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const moduleName As String = "clsExeCmd"
Private rsSO033Vod As New ADODB.Recordset
Private rsSO555C As New ADODB.Recordset
Private rsSO555D As New ADODB.Recordset
Private FCustId As String
Private FFaciSNO As String
Private FFaciSeqNo As String
Private FDeclarantName As String
Private FSO004RowId As String
Private intMessageType As Integer
Private strSendContext As String
Private lngErrTotal As Long
Private strSendVodAccountId As String
Private strSendNagraId As String
Private blnSendSMS As Boolean
Private strInsSO033VODMsg As String
Public Property Let uSmsConnectString(ByVal vData As String)
    fSmsConnectString = vData
End Property
Public Property Let uSmsTels(ByVal vData As String)
    fSmsTels = vData
End Property
Public Property Let uSmsErrCount(ByVal vData As Long)
    fSmsErrCount = vData
End Property

'設定錯誤Log的檔名(包含路徑)
Public Property Let SysPath(ByVal vData As String)
  On Error Resume Next
    strSysLog = vData
    gErrLogPath = strSysLog
End Property
'取得錯誤Log的檔名(包含路徑)
Public Property Get SysPath() As String
    On Error Resume Next
    SysPath = strSysLog
    
End Property
'設定CDRINI路徑
Public Property Let CDRIniFile(ByVal vData As String)
  On Error Resume Next
    strCDRIniFile = vData
End Property
'控制台TimeOut時間
Public Property Let SetWaitTime(ByVal vData As Long)
  On Error Resume Next
    lngCmdTimeOut = vData
End Property
'設定Connection 物件
Public Property Set uCn(ByRef vData As ADODB.Connection)
  On Error Resume Next
    Set gcnGi = vData
End Property
'設定執行物件裡Fun執行時間
Public Property Let ExeRunTime(ByVal vData As String)
  On Error Resume Next
    strRunTime = vData
    If Not IsDate(strRunTime) Then
        strRunTime = GetRightNow
    End If
End Property
Public Property Let ugaryGi(ByVal vData As Variant)
  On Error Resume Next
    Call GetGlobal(CStr(vData))
    gCompCode = garyGi(9)
End Property
'設定OwnerName
Public Property Let uOwnerName(ByVal vData As String)
  On Error Resume Next
    garyGi(16) = vData
End Property
Public Property Get uOwnerName() As String
  On Error Resume Next
    uOwnerName = garyGi(16)
End Property
Public Property Get uCompCode() As String
  On Error Resume Next
    uCompCode = garyGi(9)
End Property
Public Property Let uErrCount(ByVal vData As Long)
  On Error Resume Next
    lngErrTotal = vData
End Property
''傳送SMS至Nagra的ConnectString
'Public Property Let uSmsConnectString(ByVal vData As String)
'  On Error Resume Next
'    fSmsConnectString = vData
'End Property
'整批一次處理
Public Function ExeAllCommit(ByVal aCmdSeqNo As String, _
                        Optional ByRef strMaxIncurredDate As String = "", _
                        Optional ByVal aAddSO182 As Boolean = True, _
                        Optional ByVal aSendA3Cmd As Boolean = False, _
                        Optional ByVal aSendNagraCmd As Boolean = False, _
                        Optional ByVal aWriteIniTime As Boolean = False) As Boolean
  On Error GoTo ChkErr
    Dim blnSO033VodDouble As Boolean
    strSendVodAccountId = ""
    strSendNagraId = ""
    If Not GetSO555C(aCmdSeqNo, rsSO555C) Then Exit Function
    rsSO555C.ActiveConnection = Nothing
    If Not rsSO555C.EOF Then
        rsSO555C.MoveFirst
        Do While Not rsSO555C.EOF
            If Not InsSO033VOD(rsSO555C, , blnSO033VodDouble, False) Then
                If Not blnSO033VodDouble Then   '如果不是Double 則離開Function
                    Exit Function
                End If
            Else
'                If Len(rsSO555C("BIncurredDate") & "") > 0 Then
'                    If IsDate(rsSO555C("BIncurredDate") & "") Then
'                        If CDate(rsSO555C("BIncurredDate") & "") > CDate(FMaxIncurredDate) Then
'                            FMaxIncurredDate = rsSO555C("BIncurredDate") & ""
'                        End If
'                    End If
'                End If
            End If
            If Not blnSO033VodDouble Then   '如果沒有重複要Upd SO004G與新增SO182
                Dim aFrequencyType As String
                aFrequencyType = UCase(rsSO555C("FrequencyType") & "")
                If (aFrequencyType = "IMP") Or (aFrequencyType) = "MUL" Then
                    '更新SO004G的點數值,並新增資料至SO182
                    If Not UpdSO004G(rsSO555C("VODACCOUNTID") & "", _
                                    rsSO555C("VALUE") & "", _
                                    aSendA3Cmd, aSendNagraCmd, aAddSO182, rsSO033Vod) Then
                                    Exit Function
                    Else
                        
                    
                        '#6388 最後再呼叫Jacky的Function，計算點數 By Kin 2013/01/15
'                        If Not CalcUseCredit(rsSO555C("VODACCOUNTID") & "") Then
'                            Call WriteErr("計算點數失敗", "ExeAllCommit", "CalcUseCredit")    '寫入自訂的錯誤
'                            Exit Function
'                        End If
                        '測試不OK,改用呼叫CreditObject的方式 By Kin 2013/03/15
                        Dim clsChargeInterface As Object
                        Dim clsAlterSO003 As Object
                        Set clsChargeInterface = CreateObject("csAlterSO003.clsInterface")
                        With clsChargeInterface
                            Set .uConn = gcnGi
                            .uOwnerName = GetOwner
                            Set clsAlterSO003 = .objAction("csAlterSO003.clsAlterSO003")
                            If .uErr > 0 Then Exit Function
                        End With
                        If Not clsAlterSO003.CalcUseCredit(rsSO555C("VODACCOUNTID") & "") Then
                            Call WriteErr("計算點數失敗", "ExeAllCommit", "CalcUseCredit")    '寫入自訂的錯誤
                        End If
                        Set clsChargeInterface = Nothing
                        Set clsAlterSO003 = Nothing
                    End If
                End If
            End If
            blnSO033VodDouble = False
            rsSO555C.MoveNext
        Loop
'        If aWriteIniTime And strCDRIniFile <> "" Then
'            Call WriteCDRParaIni
'        End If
    End If
    ExeAllCommit = True
    Exit Function
ChkErr:
    Call ErrCDRSub(moduleName, "ExeAllCommit")
    
End Function
Private Function GetSO555D(ByVal aCmdSeqNo As String, ByRef aRs As Recordset) As Boolean
  On Error GoTo ChkErr
    Dim aSqL As String
    aSqL = "SELECT A.*,B.CMDSTATUS CMDSTATUS2,B.ExchTime " & _
        " FROM " & GetOwner & "SO555D A," & GetOwner & "SO555 B " & _
        " WHERE A.CMDSTATUS IN ( 'W','P') AND A.SEQNO <> " & aCmdSeqNo & _
        " AND A.SEQNO=B.SEQNO " & _
        " AND A.ERRORCOUNT <= " & fSmsErrCount & _
        " AND B.CMD='Z3' " & _
        " ORDER BY A.SEQNO "
    If Not GetCDRRs(aRs, aSqL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    aRs.ActiveConnection = Nothing
    aRs.AddNew
    aRs("SEQNO") = aCmdSeqNo
    aRs("CMDSTATUS") = "S"
    aRs("CMDSTATUS2") = "S"
    aRs("ERRORCOUNT") = 0
    aRs("ExchTime") = Now
    aRs.Update
    GetSO555D = True
  Exit Function
ChkErr:
Call ErrCDRSub(moduleName, "GetSO555D")
End Function

'一次處理一筆..並逐筆Commit
Public Function ExeOneCommit(ByVal aCmdSeqNo As String, _
                            ByRef lngErrCount As Long, _
                            ByRef strMaxIncurredDate As String, _
                            Optional AddSO182 As Boolean = True, _
                            Optional SendA3Cmd As Boolean = True, _
                            Optional SendNagraCmd As Boolean = True, _
                            Optional WriteIniTime As Boolean = True) As Boolean
  On Error GoTo ChkErr
    ExeOneCommit = False
    InitValue
    Dim blnSO033VodDouble As Boolean
    Dim blnBeginTrans As Boolean
    Dim aErrTime As Date
    aErrTime = Now
    strSendVodAccountId = ""
    blnSendSMS = False
    '#6242 增加抓取前一次連線逾時的資料 By Kin 2012/04/02
    If Not GetSO555D(aCmdSeqNo, rsSO555D) Then Exit Function
    If Not rsSO555D.EOF Then rsSO555D.MoveFirst
    '#6242 處理前次未處理完成的資料 By Kin 2012/04/02
    Do While Not rsSO555D.EOF
        aCmdSeqNo = rsSO555D("SeqNo")
        Select Case UCase(rsSO555D("CmdStatus2"))
            Case "E", "T"
                 gcnGi.Execute "UPDATE " & GetOwner & "SO555D SET " & _
                        " CMDSTATUS='" & rsSO555D("CmdStatus2") & "' " & _
                        " WHERE SEQNO=" & aCmdSeqNo
                If CDate(rsSO555D("ExchTime")) < aErrTime Then
                    aErrTime = rsSO555D("ExchTime")
                End If
                blnSendSMS = True
                SendSms aCmdSeqNo
                GoTo lNext
            Case "S"
                gcnGi.Execute "DELETE " & GetOwner & "SO555D WHERE SEQNO=" & aCmdSeqNo
            Case Else
                gcnGi.Execute "UPDATE " & GetOwner & "SO555D SET " & _
                        " ERRORCOUNT=ERRORCOUNT+1,CMDSTATUS='" & rsSO555D("CMDSTATUS2") & "' " & _
                        " WHERE SEQNO=" & aCmdSeqNo
                If Val(rsSO555D("ErrorCount") & "") + 1 > fSmsErrCount Then
                    blnSendSMS = True
                    SendSms aCmdSeqNo
                End If
                GoTo lNext
        End Select
        
        If Not GetSO555C(aCmdSeqNo, rsSO555C) Then Exit Function
        rsSO555C.ActiveConnection = Nothing
        If Not rsSO555C.EOF Then
            rsSO555C.MoveFirst
            Do While Not rsSO555C.EOF
                strInsSO033VODMsg = Empty
                gcnGi.BeginTrans
                blnBeginTrans = True
                If Not InsSO033VOD(rsSO555C, , blnSO033VodDouble) Then
                    If Not blnSO033VodDouble Then   '如果不是Double 則代表真的新增有錯誤 離開Function
                        lngErrCount = lngErrCount + 1
                        gcnGi.RollbackTrans
                        
                        If lngErrCount >= lngErrTotal And lngErrTotal > 0 Then
                            Exit Function
                        Else
                            GoTo lDo
                        End If
                    End If
                Else
                    If Len(rsSO555C("BIncurredDate") & "") > 0 Then
                        If IsDate(rsSO555C("BIncurredDate") & "") Then
                            If CDate(rsSO555C("BIncurredDate") & "") > CDate(FMaxIncurredDate) Then
                                FMaxIncurredDate = rsSO555C("BIncurredDate") & ""
                            End If
                        End If
                    End If
                End If
                If Not blnSO033VodDouble Then   '如果沒有重複要Upd SO004G與新增SO182
                    Dim aFrequencyType As String
                    aFrequencyType = UCase(rsSO555C("FrequencyType") & "")
                    If (aFrequencyType = "IMP") Or (aFrequencyType) = "MUL" Then
                        '更新SO004G的點數值,並新增資料至SO182
                        If Not UpdSO004G(rsSO555C("VODACCOUNTID") & "", _
                                        rsSO555C("VALUE") & "", _
                                        SendA3Cmd, SendNagraCmd, AddSO182, rsSO033Vod) Then
                            lngErrCount = lngErrCount + 1
                            gcnGi.RollbackTrans
                            If lngErrCount >= lngErrTotal And lngErrTotal > 0 Then
                                Exit Function
                            Else
                                GoTo lDo
                            End If
                        Else
                            '#6388 最後再呼叫Jacky的Function，計算點數 By Kin 2013/01/15
'                            If Not CalcUseCredit(rsSO555C("VODACCOUNTID") & "") Then
'                                Call WriteErr("計算點數失敗", "ExeOneCommit", "CalcUseCredit")     '寫入自訂的錯誤
'                                Exit Function
'                            End If
                            '測試不OK,改用呼叫CreditObject的方式 By Kin 2013/03/15
                            Dim clsChargeInterface As Object
                            Dim clsAlterSO003 As Object
                            Set clsChargeInterface = CreateObject("csAlterSO003.clsInterface")
                            With clsChargeInterface
                                Set .uConn = gcnGi
                                .uOwnerName = GetOwner
                                Set clsAlterSO003 = .objAction("csAlterSO003.clsAlterSO003")
                                If .uErr > 0 Then Exit Function
                            End With
                            If Not clsAlterSO003.CalcUseCredit(rsSO555C("VODACCOUNTID") & "") Then
                                Call WriteErr("計算點數失敗", "ExeAllCommit", "CalcUseCredit")    '寫入自訂的錯誤
                            End If
                            Set clsChargeInterface = Nothing
                            Set clsAlterSO003 = Nothing
                        End If
                    End If
                End If
                gcnGi.CommitTrans
                blnBeginTrans = False
                blnSO033VodDouble = False
lDo:
                '#6242 要記錄為什麼新增至SO033VOD為什麼不成功 By Kin 2012/04/03
                If strInsSO033VODMsg <> Empty Then
                    gcnGi.Execute "UPDATE " & GetOwner & "SO555C SET ErrMsg='" & strInsSO033VODMsg & "' " & _
                        " WHERE ROWID='" & rsSO555C("ROWID") & "'"
                End If
                rsSO555C.MoveNext
            Loop
            '如果都沒有錯誤將日期寫入INI檔
            If WriteIniTime And strCDRIniFile <> "" Then
                If lngErrCount = 0 Then Call WriteCDRParaIni
            End If
        End If
lNext:
        rsSO555D.MoveNext
    Loop
    '#6242 如果需要發送簡訊，則要把INI檔的時間恢復原錯誤時間 By Kin 2012/04/02
    If blnSendSMS Then
        If aErrTime < CDate(FMaxIncurredDate) Then
            FMaxIncurredDate = CStr(aErrTime)
            WriteCDRParaIni
        End If
    End If
    ExeOneCommit = True
    Exit Function
ChkErr:
  If blnBeginTrans Then gcnGi.RollbackTrans
  Call ErrCDRSub(moduleName, "ExeOnCommit")
End Function
Private Function FmtDate(ByVal aValue As String) As String
  On Error GoTo ChkErr
    Dim strTmp As String
    If IsDate(aValue) Then FmtDate = aValue: Exit Function
    strTmp = Left(aValue, 4) & "/" & Mid(aValue, 5, 2) & _
              "/" & Mid(aValue, 7, 2) & " " & _
              Mid(aValue, 9, 2) & ":" & _
              Mid(aValue, 11, 2) & ":" & _
              Right(aValue, 2)
    FmtDate = strTmp
    Exit Function
ChkErr:
  FmtDate = ""
End Function
Public Function SendZ2Cmd(ByVal strBillStartDate As String, _
                            ByVal strBillEndDate As String, _
                            ByRef rs004 As ADODB.Recordset, _
                            Optional ByRef RetCmdSeqNo As String = "", _
                            Optional ByRef RetMsg As String = "") As Boolean
                            
  On Error GoTo ChkErr
    Dim obj As Object
    Dim blnRet As Boolean
    Dim aMsg As String
    strVODCmdSeqNo = "-1"
    strBillStartDate = FmtDate(strBillStartDate)
    strBillEndDate = FmtDate(strBillEndDate)
    If Not IsDate(strBillStartDate) Then
        RetMsg = "傳送Z2命令起始日期格式不正確！傳送值為:" & strBillStartDate
        Call WriteErr("傳送Z2命令起始日期格式不正確！傳送值為:" & strBillStartDate, moduleName, "SendZ2Cmd")
        Exit Function
    End If
    If Not IsDate(strBillEndDate) Then
        RetMsg = "傳送Z2命令終止日期格式不正確！傳送值為:" & strBillStartDate
        Call WriteErr("傳送Z2命令終止日期格式不正確！傳送值為:" & strBillStartDate, moduleName, "SendZ2Cmd")
        Exit Function
    End If
    Set obj = CreateObject("csVODcontrol.clsExeCommand")
    blnRet = obj.ExeVodCommand(gcnGi, rs004, "Z2", garyGi, , , strRunTime, , _
                , , , , , strBillStartDate, _
                strBillEndDate, , , , , aMsg, lngCmdTimeOut, strVODCmdSeqNo)
    
    If Not blnRet Then
        If aMsg = "" Then aMsg = "未知錯誤"
        RetMsg = aMsg
        Call WriteErr("傳送Z2命令錯誤！ " & aMsg, moduleName, "SendZ2Cmd")
        Exit Function
    End If
    If strVODCmdSeqNo < 0 Then
        aMsg = "取得SO555 SEQNO錯誤！"
        RetMsg = aMsg
        Call WriteErr(aMsg, moduleName, "SendZ2Cmd")
    End If
    RetCmdSeqNo = strVODCmdSeqNo
    SendZ2Cmd = blnRet
    On Error Resume Next
    'Call CloseRecordset(rsVir)
    Set obj = Nothing
    Exit Function
ChkErr:
    Set obj = Nothing
    Call ErrCDRSub(moduleName, "SendZ2Cmd")
End Function
'傳送Z3的命令到VOD控制台
Public Function SendZ3Cmd(ByVal strBillStartDate As String, _
                            Optional ByRef RetCmdSeqNo As String = "", _
                            Optional ByVal ReadINITime As Boolean = False) As Boolean
  On Error GoTo ChkErr
    Dim obj As Object
    Dim blnRet As Boolean
    Dim rsVir As New ADODB.Recordset
    Dim strCD088SQL As String
    Dim aMsg As String
    strVODCmdSeqNo = "-1"
    FMaxIncurredDate = ""
    SetRsField rsVir    '設定一個虛擬的,以便可以傳給控制台
    If ReadINITime Then     '判斷是否自動讀取INI檔的傳送時間
        If strCDRIniFile <> "" Then
            strBillStartDate = GetSendTime
            If strBillStartDate <> "" Then
                strBillStartDate = FmtDate(strBillStartDate)
            End If
        Else
            Call WriteErr("未給予INI檔名稱與路徑", moduleName, "SendZ3Cmd")
            Exit Function
        End If
    End If
     
    strBillStartDate = FmtDate(strBillStartDate)
    If Not IsDate(strBillStartDate) Then
        Call WriteErr("傳送Z3命令日期格式不正確！傳送值為:" & strBillStartDate, moduleName, "SendZ3Cmd")
        Exit Function
    Else
        FMaxIncurredDate = strBillStartDate '記錄傳送日期
        strBillStartDate = DateAdd("n", -1, CDate(strBillStartDate))
    End If
    
    Set obj = CreateObject("csVODcontrol.clsExeCommand")
    blnRet = obj.ExeVodCommand(gcnGi, rsVir, "Z3", garyGi, , , strRunTime, , _
                , , , , , strBillStartDate, _
                GetRightNow, , , , , aMsg, lngCmdTimeOut, strVODCmdSeqNo)
    
    If Not blnRet Then
        If aMsg = "" Then aMsg = "未知錯誤"
        Call WriteErr("傳送Z3命令錯誤！ " & aMsg, moduleName, "SendZ3Cmd")
        '#6242 因為還要抓取SO555D的資料,所以要回傳True
        'Exit Function
    End If
    '#6242 因為還要抓取SO555D的資料,所以要回傳True
'    If strVODCmdSeqNo < 0 Then
'        aMsg = "取得SO555 SEQNO錯誤！"
'        Call WriteErr(aMsg, moduleName, "SendZ3Cmd")
'        SendZ3Cmd = False
'        Exit Function
'    End If
    RetCmdSeqNo = strVODCmdSeqNo
    'SendZ3Cmd = blnRet
    SendZ3Cmd = True
    On Error Resume Next
    Call CloseRecordset(rsVir)
    Set obj = Nothing
    Exit Function
ChkErr:
    Set obj = Nothing
    Call ErrCDRSub(moduleName, "SendZ3Cmd")
End Function
'設定虛擬Recordset,給Z3命令使用的
Private Sub SetRsField(ByRef rs As ADODB.Recordset)
  On Error GoTo ChkErr
    With rs
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Fields.Append "CustId", adBSTR, 10, adFldIsNullable
        .Fields.Append "FaciSNo", adBSTR, 10, adFldIsNullable
        .Fields.Append "SmartCardNo", adBSTR, 15, adFldIsNullable
        .Fields.Append "VODAccountId", adBSTR, 20, adFldIsNullable
        .Fields.Append "CompCode", adBSTR, 10, adFldIsNullable
    End With
    rs.Open
    rs.AddNew
    rs("CompCode") = gCompCode
    rs.Update
    Exit Sub
ChkErr:
  Call ErrCDRSub(moduleName, "SetRsField")
End Sub
'傳送Z3命令後,撈取SO555C的資料 aCmdSeqNo--Z3傳送回來SO555的SeqNo、rs555C--回傳撈取到的資料
Public Function GetSO555C(ByVal aCmdSeqNo As String, ByRef rs555C As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim strSQL2 As String
    Dim aMsg As String
    If Val(aCmdSeqNo) = 0 Then
        aMsg = "SEQNO為空值"
        Call WriteErr(aMsg, moduleName, "GetSO555C")    '寫入自訂的錯誤
        Exit Function
    End If
    strSQL = "Select A.*,A.ROWID From " & GetOwner & "SO555C A Where SeqNo = " & Val(aCmdSeqNo) & _
            " ORDER BY VODACCOUNTID "
            
            
'    strSQL2 = "Select A.*,1 NType From " & GetOwner & "SO555C A," & GetOwner & "SO555D B " & _
'                " Where B.SeqNo = A.SeqNo AND B.CmdStatus IN ('P','W') AND B.SEQNO <> " & Val(aCmdSeqNo)
'
'    strSQL = "SELECT * FROM (" & strSQL & " UNION ALL " & strSQL2 & ") ORDER BY NTYPE,VODACCOUNTID"
    If Not GetCDRRs(rs555C, strSQL, gcnGi, adUseClient, _
                    adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not GetCreditType Then Exit Function
    GetSO555C = True
    Exit Function
ChkErr:
  Call ErrCDRSub(moduleName, "GetSO555C")
End Function
'取得點數種類的代碼與名稱
Private Function GetCreditType() As Boolean
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    If Not GetCDRRs(rsTmp, "Select CodeNo,Description From " & GetOwner & "CD108" & _
                      " WHERE REFNO = 3 AND NVL(STOPFLAG,0)=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsTmp.EOF Then
        strCreditTypeCode = rsTmp(0) & ""
        strCreditTypeName = rsTmp(1) & ""
    End If
    GetCreditType = True
    On Error Resume Next
    Call CloseRecordset(rsTmp)
    Exit Function
ChkErr:
  Call ErrCDRSub(moduleName, "GetCreditType")
End Function
'以SO555C的資料新增至rsSO033VOD
Public Function InsSO033VOD(ByRef rs555C As ADODB.Recordset, _
                            Optional ByRef rs033VodRet As ADODB.Recordset = Nothing, _
                            Optional ByRef blnDouble As Boolean, _
                            Optional ByVal blnPrDate As Boolean = True) As Boolean
                            
  On Error GoTo ChkErr
    If Not rs555C.EOF Then
        If Not GetCDRRs(rsSO033Vod, "Select * From " & GetOwner & "SO033VOD " & _
                    " WHERE 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        
        Set rs033VodRet = rsSO033Vod
        If Not GetSO004Data(rs555C("VODAccountID") & "", blnPrDate) Then
            strInsSO033VODMsg = "請檢查VODAccountId是否存在或是否為拆機"
            Exit Function
        End If
        '先檢查SO033VOD裡面是否已有資料,如果沒有才做新增的動作
        If gcnGi.Execute("Select Count(CustId) From " & GetOwner & "SO033Vod " & _
                    " WHERE VODUID = " & rs555C("VODUID"))(0) <= 0 Then
            With rsSO033Vod
                .AddNew
                .Fields("CompCode") = gCompCode
                .Fields("CustId") = FCustId
                .Fields("VODAccountID") = rs555C("VODAccountID") & ""
                .Fields("itemUID") = rs555C("itemUID") & ""
                .Fields("itemName") = rs555C("itemName") & ""
                If IsDate(rs555C("BincurredDate") & "") Then
                    .Fields("incurredDate") = rs555C("BincurredDate")
                End If
                .Fields("Value") = Val(rs555C("Value") & "")
                .Fields("FrequencyType") = rs555C("FrequencyType") & ""
                .Fields("ItemType") = rs555C("ItemType") & ""
                .Fields("SeqNo") = Get_SO033VOD_Seq
                .Fields("VODUID") = rs555C("VODUID")
                .Update
            End With
        Else
            InsSO033VOD = False
            blnDouble = True
            strInsSO033VODMsg = "SO033VOD.VODUID已存在，VODUID:" & rs555C("VODUID")
            Exit Function
        End If
        
    End If
    InsSO033VOD = True
    Exit Function
ChkErr:
  Call ErrCDRSub(moduleName, "InsSO033Vod")
End Function
Private Sub Class_Initialize()
'  On Error Resume Next
'    '給予內定的錯誤訊息Log路徑,如果有傳再以傳入的為主
'    Dim sPath As String
'    sPath = App.Path
'    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
'    sPath = sPath & "SysErr.txt"
'    strErrPath = sPath
'    gErrLogPath = sPath
'    '內定傳送控制台TimeOut時間
'    lngCmdTimeOut = Val(ReadINIData(strCDRIniFile, "CDRCenter", "CmdTimeOut"))
End Sub
Private Sub InitValue()
  On Error Resume Next
    '給予內定的錯誤訊息Log路徑,如果有傳再以傳入的為主
    Dim sPath As String
    sPath = App.Path
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    sPath = sPath & "SysErr.txt"
    strErrPath = sPath
    gErrLogPath = sPath
    '內定傳送控制台TimeOut時間
    lngCmdTimeOut = Val(ReadINIData(strCDRIniFile, "CDRCenter", "CmdTimeOut"))
    Dim test As String
    test = ReadINIData(strCDRIniFile, "SMSTEL", Empty)
End Sub
'取出SO004的一些基本資料
Private Function GetSO004Data(ByVal aVodAccountId As String, ByVal IsPrDate As Boolean) As Boolean
  On Error GoTo ChkErr
    Dim rsSO004 As New ADODB.Recordset
    If Not GetCDRRs(rsSO004, "SELECT A.RowId,A.* FROM " & GetOwner & "SO004 A " & _
                    " WHERE VODAccountId=" & aVodAccountId & IIf(IsPrDate, " AND PRDATE IS NULL ", ""), _
                    gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    FCustId = Empty
    FFaciSNO = Empty
    FFaciSeqNo = Empty
    FDeclarantName = Empty
    FSO004RowId = Empty
    If Not rsSO004.EOF Then
        FCustId = rsSO004("CustID") & ""
        FFaciSNO = rsSO004("FaciSNO") & ""
        FFaciSeqNo = rsSO004("SeqNo") & ""
        FDeclarantName = rsSO004("DeclarantName") & ""
        FSO004RowId = rsSO004("RowId") & ""
    Else
        Exit Function
    End If
    GetSO004Data = True
    On Error Resume Next
    Call CloseRecordset(rsSO004)
    Exit Function
ChkErr:
  Call ErrCDRSub(moduleName, "GetSO004Data")
End Function
Private Function Get_SO182_Seq() As String
  On Error GoTo ChkErr
    Get_SO182_Seq = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO182_SEQNO.NEXTVAL FROM DUAL").GetString & "")
    Exit Function
ChkErr:
  Call ErrSub(moduleName, "Get_SO182_Seq")
End Function
Private Function Get_INV037_CMDSEQNO(ByRef cnCom As ADODB.Connection) As String
On Error GoTo ChkErr
    Get_INV037_CMDSEQNO = RPxx(cnCom.Execute("SELECT S_INV037_CMDSEQNO.NEXTVAL FROM DUAL").GetString & "")
  Exit Function
ChkErr:
    Call ErrSub(moduleName, "Get_INV037_CMDSEQNO")

End Function
Private Function Get_SO033VOD_Seq() As String
On Error GoTo ChkErr
    Get_SO033VOD_Seq = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO033VOD_SEQNO.NEXTVAL FROM DUAL").GetString & "")
    Get_SO033VOD_Seq = Format(GetRightNow, "yyyymmdd") & _
                        Right("000" & FCustId, 3) & _
                        Format(Get_SO033VOD_Seq, "000000000")
  Exit Function
ChkErr:
    Call ErrSub(moduleName, "Get_SO033VOD_Seq")
End Function

'新增資料至SO182 aMinusType:新增的種類 aValue:扣除點數
Public Function InsSO182(ByVal aMinusType As Integer, ByVal aValue As String, _
                            ByRef rs033VOD As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr:
    Dim strSQL As String
    Dim rsSO182 As New ADODB.Recordset
    If Not GetCDRRs(rsSO182, "Select * From " & GetOwner & "SO182 Where 1 = 0", _
                gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    With rsSO182
        .AddNew
        .Fields("CompCode") = gCompCode
        .Fields("ServiceType") = "D"
        .Fields("RecTime") = rs033VOD("IncurredDate")
        .Fields("RecEn") = "EPG"
        .Fields("FaciSeqNo") = FFaciSeqNo
        .Fields("UpdEn") = "EPG"
        .Fields("UpdTime") = Format(strRunTime, "EE/MM/DD HH:MM:SS")
        .Fields("MinusType") = aMinusType
        .Fields("CreditTypeCode") = IIf(strCreditTypeCode <> "", strCreditTypeCode, Null)
        .Fields("CreditTypeName") = IIf(strCreditTypeName <> "", strCreditTypeName, Null)
        .Fields("MinusCredit") = aValue
        .Fields("VODAccountID") = rs033VOD("VODACCOUNTID")
        .Fields("CreditSeqNo") = rs033VOD("SEQNO")
        .Fields("ModifyReason") = rs033VOD("itemName") & ""
        .Fields("SeqNo") = Get_SO182_Seq
        .Fields("STOPFLAG") = 0
        .Update
    End With
    InsSO182 = True
    On Error Resume Next
    Call CloseRecordset(rsSO182)
    Exit Function
ChkErr:
  Call ErrCDRSub(moduleName, "InsSO182")
End Function
'回傳要Upd SO004G的點數值--aPointValue(SO004G的值)、aValue(要扣的點數值)、aRetValue未扣完的點數
Private Function CalculPoint(ByVal aPointValue As Long, _
                            ByVal aValue As Long, ByRef aRetValue As Long) As Long
    Dim lngAnsValue As Long
    lngAnsValue = aPointValue - aValue
    If lngAnsValue >= 0 Then
        aRetValue = 0
        CalculPoint = lngAnsValue
    Else
        aRetValue = Abs(lngAnsValue)
        CalculPoint = 0
    End If
End Function
Private Sub Class_Terminate()
  On Error Resume Next
    FCustId = Empty
    FFaciSNO = Empty
    FFaciSeqNo = Empty
    FDeclarantName = Empty
    FSO004RowId = Empty
    strVODCmdSeqNo = -1
    strCDRIniFile = ""
    strCreditTypeCode = ""
    strCreditTypeName = ""
    lngCmdTimeOut = 0
    Call CloseRecordset(rsSO033Vod)
    Call CloseRecordset(rsSO555C)
End Sub

'更新SO004G的資料,並新增資料至SO182
Public Function UpdSO004G(ByVal aVodAccount As String, _
                            ByVal aValue As String, _
                            Optional ByVal blnSendA3Cmd As Boolean = True, _
                            Optional ByVal blnSendNagraCmd As Boolean = True, _
                            Optional ByVal blnAddSO182 As Boolean = True, _
                            Optional ByRef rs033VOD As ADODB.Recordset = Nothing) As Boolean
 On Error GoTo ChkErr
    Dim strQry As String
    Dim rsSO004G As New ADODB.Recordset
    Dim lngRetValue As Long
    Dim lngUpdValue As Long
    Dim lngMinusValue1 As Long
    Dim lngMinusValue2 As Long
    Dim lngMinusValue3 As Long
    Dim aSO182Seq As String
    '如果要自動新增到SO182,則rs033VOD不能為Nothing
    If (blnAddSO182) And (rs033VOD Is Nothing) Then
        Exit Function
    End If
    strQry = "SELECT * FROM " & GetOwner & "SO004G " & _
                " WHERE VODAccountID='" & aVodAccount & "'"
    If Not GetCDRRs(rsSO004G, strQry, gcnGi, adUseClient, _
                    adOpenKeyset, adLockOptimistic) Then Exit Function
    aSO182Seq = Empty
    lngMinusValue1 = 0
    lngMinusValue2 = 0
    lngMinusValue3 = 0
    If Not rsSO004G.EOF Then
    
        '***********************************新增儲值至SO182*******************************
        '如果儲存點數有值,先扣掉儲存點數
        If Val(rsSO004G("PrePay") & "") > 0 Then
            lngRetValue = 0
            lngUpdValue = CalculPoint(Val(rsSO004G("PrePay") & ""), aValue, lngRetValue)
            lngMinusValue1 = Val(rsSO004G("PrePay") & "") - lngUpdValue   '原本的值-要Upd的值=扣掉多少點(要新增至SO182用)
            aValue = lngRetValue
            rsSO004G("PrePay").Value = lngUpdValue
            '最後再新增就好,因為要合併成一筆 For Jacky By Kin 2013/03/20
'            If blnAddSO182 Then '新增一筆扣款種煩為 "儲值" 的到SO182
'                If Not InsSO182(1, lngMinusValue, rs033VOD) Then Exit Function
'            End If
        Else
            lngRetValue = aValue
        End If
        
        
        
        '***********************************新增贈送至SO182*******************************
        '如果贈送點數有值,再扣掉贈送點數
        If Val(rsSO004G("Present") & "") > 0 Then
            If (lngRetValue > 0) Then
                aValue = lngRetValue
                lngRetValue = 0
                lngMinusValue2 = 0
                lngUpdValue = CalculPoint(Val(rsSO004G("Present") & ""), aValue, lngRetValue)
                lngMinusValue2 = Val(rsSO004G("Present") & "") - lngUpdValue
                rsSO004G("Present") = lngUpdValue
                 '最後再新增就好,因為要合併成一筆 For Jacky By Kin 2013/03/20
'                If blnAddSO182 Then '新增一筆扣款種類為贈送的到SO182
'                    If Not InsSO182(2, lngMinusValue, rs033VOD) Then Exit Function
'                End If
            End If
        Else
            
            lngRetValue = aValue
        End If
        
        
        '***********************************新增預借至SO182*******************************
        '不夠扣要加在未付點數
        If lngRetValue > 0 Then
            lngMinusValue3 = Abs(lngRetValue)
            rsSO004G("NoPayCredit") = Val(rsSO004G("NoPayCredit") & "") + Abs(lngRetValue)
'            If blnAddSO182 Then '新增一筆扣款種類為預借的到SO182
'                If Not InsSO182(3, lngMinusValue, rs033VOD) Then Exit Function
'            End If
        End If
        
        '將金額加總新增一筆SO182 For Jacky By Kin 2013/03/20
        If blnAddSO182 Then '新增一筆扣款種類為預借的到SO182
            If Not InsSO182(3, lngMinusValue1 + lngMinusValue2 + lngMinusValue3, rs033VOD) Then Exit Function
        End If
        
        
        '************************************傳送A3(暫停命令)至VOD控制台*********************************
        '未付費點數>=預借點數 要傳送A3命令給控制台
        If blnSendA3Cmd Then
            If Val(rsSO004G("NoPayCredit") & "") >= Val(rsSO004G("VODCredit") & "") Then
                 '判斷同一個VodAccountId是否有送過，有送過就不要再送了 By Kin 2009/12/02
                If strSendVodAccountId <> aVodAccount Then
                    strSendVodAccountId = aVodAccount
                    '#5667 A3命令要忽略錯誤 By Kin 2010/05/27
                    'If Not SendA3Cmd(aVodAccount) Then Exit Function
                    SendA3Cmd aVodAccount
                    '#5409 如果有傳送A3命令，要將StopType更改為1 By Kin 2009/12/02
                    rsSO004G("StopType").Value = 1
                    
                End If
            End If
        End If
        'strSendVodAccountId = aVodAccount
        
        
        '************************************傳送簡訊至Nagra*********************************************
        '是否要傳送簡訊至Nagra
        If blnSendNagraCmd Then
            '判斷SO041是否啟動傳送機制
            If gcnGi.Execute("Select Nvl(SendMessage,0) From " & GetOwner & "SO041 " & _
                            " Where SysId=" & gCompCode)(0) = 1 Then
                Dim lngSparePay As Long
                Dim lngAlarmPay As Long
                
                '計算剩餘點數=信用額度-未付費點數
                lngSparePay = Val(rsSO004G("VODCredit") & "") - Val(rsSO004G("NoPayCredit") & "")
                '取出警戒值
                lngAlarmPay = Val(gcnGi.Execute("Select Nvl(VODLowValue,0) From " & GetOwner & "SO041 " & _
                            " Where SysId=" & gCompCode)(0) & "")
                If lngSparePay <= lngAlarmPay Then  '剩餘點數小於SO041的警戒值
                    Dim lngSendCount As Long
                    '取出最多傳送幾次簡訊
                    lngSendCount = Val(gcnGi.Execute("Select Nvl(MaxSendTime,0) From " & GetOwner & "SO041 " & _
                                " Where SysId=" & gCompCode)(0) & "")
                    '判斷發送次數是否已經超過SO041發送次數了
                    '阿貴說如果是0代表沒有限制
                    If lngSendCount = 0 Then
                        If strSendNagraId <> aVodAccount Then
                            strSendNagraId = aVodAccount
                            Call SendNagraSMSCmd(aVodAccount)
                            rsSO004G("SendCount") = Val(rsSO004G("SendCount") & "") + 1
                        End If
                    Else
                        If strSendNagraId <> aVodAccount Then
                            
                            If Val(rsSO004G("SendCount") & "") < lngSendCount Then
                                strSendNagraId = aVodAccount
                                Call SendNagraSMSCmd(aVodAccount)
                                rsSO004G("SendCount") = Val(rsSO004G("SendCount") & "") + 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        rsSO004G.Update
    End If
    UpdSO004G = True
    Exit Function
ChkErr:
  Call ErrCDRSub(moduleName, "UpdSO004G")
End Function
Private Function SendSms(ByVal aSeqNo As String) As Boolean
  On Error GoTo ChkErr
    Dim cnSms As New ADODB.Connection
    Dim rsInv037 As New ADODB.Recordset
    Dim strSplit() As String
    Dim i As Long
    If Len(fSmsConnectString & "") = 0 And Len(fSmsTels & "") = 0 Then
        SendSms = True
        GoTo Fin
    End If
    If Not IsSMSConnect(cnSms) Then
        WriteErr "簡訊資料庫無法連接！", moduleName, "SendSms"
        SendSms = True
        GoTo Fin
    End If
    strSplit = Split(fSmsTels, ",")
    If Not GetCDRRs(rsInv037, "SELECT * FROM INV037 WHERE 1=0", cnSms, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    For i = LBound(strSplit) To UBound(strSplit)
        cnSms.BeginTrans
        rsInv037.AddNew
        rsInv037("IdentifyId1") = "1"
        rsInv037("IdentifyId2") = 0
        rsInv037("CompID") = gCompCode
        rsInv037("CONTMOBILE") = strSplit(i)
        rsInv037("INVID") = aSeqNo
        rsInv037("DATATIME") = Now
        rsInv037("DATAEN") = "Auto-CDR"
        rsInv037("CMDSEQNO") = Get_INV037_CMDSEQNO(cnSms)
        rsInv037("CONTENT") = "VOD取帳指令失敗，SEQNO=" & aSeqNo
        rsInv037.Update
        cnSms.CommitTrans
    Next i
    SendSms = True
Fin:
  On Error Resume Next
    rsInv037.Close
    Set rsInv037 = Nothing
    cnSms.Close
    Set cnSms = Nothing
  Exit Function
  
ChkErr:
  Call ErrCDRSub(moduleName, "SendSms")
End Function


Public Function SendA3Cmd(ByVal aVodAccountId As String) As Boolean
  On Error GoTo ChkErr
    Dim blnRet As Boolean
    Dim rsVir As New ADODB.Recordset
    Dim strQry As String
    Dim obj As Object
    Dim strRetErrMsg As String
    Dim blnSend As Boolean
    Set obj = CreateObject("csVODcontrol.clsExeCommand")
    strQry = "SELECT A.ROWID,A.* FROM " & GetOwner & "SO004 A " & _
          " WHERE A.VODACCOUNTID = " & aVodAccountId
          
    If Not GetCDRRs(rsVir, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsVir.EOF Then
        Call WriteErr("傳送A3命令找不到SO004資料", moduleName, "SendA3Cmd")
        Exit Function
    End If
    '#5469 如果有2筆以上的資料要以VODSTATUS=1那筆為主 By Kin 2010/01/07
    blnSend = True
    If rsVir.RecordCount > 1 Then
        blnSend = False
        rsVir.MoveFirst
        Do While Not rsVir.EOF
            If Val(rsVir("VODSTATUS") & "") = 1 Then
                blnSend = True
                Exit Do
            End If
            rsVir.MoveNext
        Loop
    End If
    If Not blnSend Then
        Call WriteErr("傳送A3命令找不到SO004資料", moduleName, "SendA3Cmd")
        Exit Function
    End If
    blnRet = obj.ExeVodCommand(gcnGi, rsVir, "A3", garyGi(), _
                                , , strRunTime, , , , , , , , , , , , , strRetErrMsg)
    If Not blnRet Then
        Call WriteErr(strRetErrMsg, moduleName, "SendA3Cmd")
        Exit Function
    End If
88:
    SendA3Cmd = blnRet
    On Error Resume Next
    Set obj = Nothing
    Call CloseRecordset(rsVir)
    Exit Function
  
ChkErr:
  Call ErrCDRSub(moduleName, "SendA3Cmd")
End Function
'傳送簡訊至Nagra命令
'cn As ADODB.Connection                         '連線參數
'rs As ADODB.Recordset                          '設備資料(SO004 or PrSO004)
'ByVal strCompCode As String                    '公司別代碼
'ByVal strProcessType As String,                '傳送的命令
'ByRef strGaryGi As String                      '傳送 garyGi 陣列使用
'ByRef strRowid As String,                      '要回傳Send_Nagra.ROWID
'Optional ByVal strResvTime As String,          '預約時間 EX:20080101235900
'Optional ByVal strChNum As String,             '選擇的頻道號碼(CD024.CodeNo 每個頻道中間用逗號隔開)
'Optional ByVal strChStopDate As String,        '頻道截止日 YYYY/MM/DD HH:MM:SS
'Optional ByVal strFaciCode As String,          '設備代碼
'Optional ByVal strFaciSNo As String,           '設備序號
'Optional ByVal strSNO As String = Empty,       '設備SNO
'Optional ByVal strOrderNo As String = Empty,   '訂單號碼
'Optional ByRef strErrMsg As String,            '回傳錯誤訊息
'Optional ByRef strMsgBack As String            '回傳訊息
'Optional ByVal intPriotity as Integer=0        '傳送訊息等級  0: noraml priority 1: high priority;
'Optional ByVal strSendMsg as String            '傳送訊息
'Optional ByVal blnOpenTmpCh as Boolean=False   '是否為試看頻道 False=不是 True=是 試看頻道
'Optional rsch As Recordset= Nothing            '傳送頻道的Recordset VARCHAR2.Fields("STBSNo"), VARCHAR2.Fields("SmartCardNo"), VARCHAR2.Fields("ChCode"), VARCHAR2.Fields("ChName"), VARCHAR2.Fields("AuthorStartDate"), VARCHAR2.Fields("AuthorStopDate")
'Optional ByRef strSeqno As String = ""         '原本回傳SOAC0202的RowID 改為回傳 CA SEQNO
'Optional ByVal strBillNo As String             '整批送命令如果有BILLNO+ITEM則填入SOAC0202或 SOAC0202TMP 內
'Optional ByVal strItem As String               '整批送命令如果有BILLNO+ITEM則填入SOAC0202或 SOAC0202TMP 內
Public Function SendNagraSMSCmd(ByVal aVodAccountId As String, Optional ByVal strMsg As String = Empty) As Boolean

  On Error GoTo ChkErr
    Dim rs004 As New ADODB.Recordset
    Dim strQry As String
    Dim obj As Object
    Dim rsCD088 As New ADODB.Recordset
    Dim blnRet As Boolean
    Dim strRowId As String
    Dim strErrMsg As String
    strQry = "SELECT A.ROWID,A.* FROM " & GetOwner & "SO004 A Where A.VodAccountId=" & aVodAccountId
    If Not GetCDRRs(rs004, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rs004.EOF Then Exit Function
    strQry = Empty
    strRowId = rs004("RowId") & ""
    '#5476 原本取CD088.REFNO=1現在改為2 By Kin 2010/01/27
    If strMsg = Empty Then
        intMessageType = gcnGi.Execute("Select Nvl(MessageType,0) From " & GetOwner & "SO041 " & _
                                    " Where SysId=" & gCompCode)(0)
        Select Case intMessageType
            Case 0
                strQry = "Select DESCRIPTION FROM " & GetOwner & "CD088 " & _
                        " WHERE MessageType = 0 AND REFNO= 2 AND STOPFLAG <> 1 "
            Case 1
                strQry = "Select DESCRIPTION FROM " & GetOwner & "CD088 " & _
                        " WHERE MessageType = 1 AND REFNO= 2 AND STOPFLAG <> 1"
            Case Else
                strQry = Empty
        End Select
    End If
    
    If strQry <> Empty Then
        If Not GetCDRRs(rsCD088, strQry, gcnGi, adUseClient, _
                        adOpenKeyset, adLockOptimistic) Then Exit Function
        If rsCD088.EOF Then
            strMsg = Empty
        Else
            strMsg = rsCD088(0) & ""
        End If
    End If
    '#5667 Nagra簡訊改用預約的方式(系統時間+5分鐘) By Kin 2010/05/27
    Dim strResvTime As String
    strResvTime = DateAdd("n", 5, Now)
    strResvTime = Format(strResvTime, "yyyymmddhhmmss")
    If strMsg <> Empty Then
        Set obj = CreateObject("csNagraSTB3.clsExeCommand")
        blnRet = obj.ExeDVNCommand(gcnGi, rs004, CStr(gCompCode), "E2", _
                         Join(garyGi, Chr(9)), strRowId, strResvTime, , , _
                         rs004("FaciCode") & "", rs004("FaciSNo") & "", , _
                         rs004("OrderNo") & "", strErrMsg, , intMessageType, strMsg)
        
    Else
        Call WriteErr("找不到簡訊內容", moduleName, "SendNagraSMSCmd")
        Exit Function
    End If
    If Not blnRet Then
        If strErrMsg = Empty Then
            strErrMsg = "未知錯誤！"
        End If
        Call WriteErr("傳送簡訊NagraCmd命令錯誤 : " & strErrMsg, moduleName, "SendNagraSMSCmd")
        Exit Function
    End If
    On Error Resume Next
    SendNagraSMSCmd = True
    Call CloseRecordset(rs004)
    Call CloseRecordset(rsCD088)
    Set obj = Nothing
    Exit Function
ChkErr:
  Call ErrCDRSub(moduleName, "SendNagraSMSCmd")
End Function

