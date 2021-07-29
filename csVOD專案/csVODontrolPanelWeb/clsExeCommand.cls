VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExeCommand"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strCustName As String
Private lngInstAddrNo As Long
Private strInstAddress  As String
Private lngNodeNo  As Long
Private lngCircuitNo  As Long
Private strZipCode  As String
Private strCustTel As String
Public Function ExeVodCommand(ByRef cn As ADODB.Connection, ByRef rs As ADODB.Recordset, _
                             ByVal strCMD As String, ByRef GaryGi() As String, _
                             Optional ByVal strCitemCode As String = Empty, _
                             Optional ByVal strCancelCitemCode As String = Empty, _
                             Optional ByVal strUpdTime As String = Empty, _
                             Optional ByVal strUpdEn As String = Empty, _
                             Optional ByVal strProcessingDate As String = Empty, _
                             Optional ByVal strSNO As String = Empty, _
                             Optional ByVal strMediaBillNo As String = Empty, _
                             Optional ByVal strReasonCode As String = Empty, _
                             Optional ByVal strReasonName As String = Empty, _
                             Optional ByVal strBillStartDate As String = Empty, _
                             Optional ByVal strBillStopDate As String = Empty, _
                             Optional ByVal strCmdSeq As String = Empty, _
                             Optional ByRef rsNew As ADODB.Recordset = Nothing, _
                             Optional ByRef strResult As String, _
                             Optional ByRef strFaultReason As String, _
                             Optional ByRef strRetErrMsg As String, _
                             Optional ByVal lngTimeOut As Long = 0, _
                             Optional ByRef strRetCmdSeq As String, _
                             Optional ByVal blnExe As Boolean = True, _
                             Optional ByRef strSTBUId As String = Empty, _
                             Optional ByVal strICCUID As String = Empty, _
                             Optional ByVal IsCancelCommand As Boolean = False, _
                             Optional ByVal PauseTime As Integer = 0) As Boolean
    On Error Resume Next
    blnShowMsg = False
    blnNoShowMsg = True
    ExeVodCommand = False
    blnExeCommand = blnExe '判斷是否執行命令再更新SO004
    intRunRetunCount = 0
    Dim blnRet As Boolean
    Dim strArry As String
    IsCancelCmd = IsCancelCommand
    '#6981 增加傳送時間間隔
    FPausetime = PauseTime
    Dim objTempWrite As TextStream
    Set objTempWrite = objFileSystem.CreateTextFile("C:\PauseTime.Txt", True)
    objTempWrite.WriteLine ("傳入時間間隔:" & FPausetime)
    objTempWrite.Close
    objTempWrite = Nothing
    
    'Set frmSO1623A = Nothing
    strArry = Join(GaryGi, Chr(9))
    Set gcnGi = cn
    Call GetGlobal(strArry)
    If Not rsNew Is Nothing Then
        Set rs2 = rsNew
    Else
        Set rs2 = Nothing
    End If
    '#7031 增加是否要過濾B15命令資料 By Kin 2015/04/28
    Call GetClctStopDateType
    blnRecordProcedure = IsRecordProcedure  '判斷是否要記錄全部程序
    'Call WriteRecordVodProcedure("0.陣列公司別：[" & GaryGi(9) & "]，實際資料公司別：[" & rs("CompCode") & "]", True)
    Call WriteRecordVodProcedure("設備序號 : " & rs("FaciSNo"))
    Call WriteRecordVodProcedure("1.開始進入無UI程序 ")
    '******************************************************************************************************************************
    '監控中心所用，當blnExeCommand=False時，則不執行命令，直接進入命令裡的SO004更新
    If Not blnExe Then
        strCMowner = gcnGi.Execute("Select CMOwner From " & GetOwner & "SO041 Where CompCode=" & rs("CompCode"))(0) & ""
        strCmdSeqNo = strCmdSeq
        If strCMowner <> Empty Then
            strCMowner = strCMowner & "."
        End If
    End If
    'Z3命令等待的時間,不可以以SO41的為主,要特別處理
    If lngTimeOut > 0 Then
        lngCDRTimeOut = lngTimeOut
    Else
        lngCDRTimeOut = 0
    End If
    '*****************************************************************************************************************************
    If rs Is Nothing Then
        Call WriteRecordVodProcedure("-99.RecordSet Is Nothing !")
        Exit Function
    End If
    If rs.EOF Then
        Call WriteRecordVodProcedure("-98.RecordSet No Data !")
        Exit Function
    End If
    If (strCMD = "A6") Or (strCMD = "A8") Or (strCMD = "A9") Then
        If rs2 Is Nothing Then
            strRetErrMsg = "沒有傳入新值資料"
            Call WriteRecordVodProcedure("-2." & strRetErrMsg)
            Exit Function
        Else
            If rs.EOF Then
                strRetErrMsg = "沒有傳入新值資料"
                Call WriteRecordVodProcedure("-3." & strRetErrMsg)
                Exit Function
            End If
        End If
    End If
    Call WriteRecordVodProcedure("2.進入執行命令程序")
    blnRet = ExeOpenVod(rs, strCMD, strCitemCode, strCancelCitemCode, strUpdTime, strUpdEn, strSTBUId, _
                                strProcessingDate, strSNO, strMediaBillNo, strReasonCode, _
                                strReasonName, IIf(rs2 Is Nothing, Nothing, rs2), strResult, strFaultReason, strRetErrMsg, _
                                strBillStartDate, strBillStopDate, strICCUID)
    Call WriteRecordVodProcedure("3.完成執行命令程序並離開", False)
    Call WriteRecordVodProcedure("4.陣列公司別：[" & GaryGi(9) & "]，實際資料公司別：[" & rs("CompCode") & "]", True)
    If blnRet Then
        strRetCmdSeq = strRetCmdSeqNo
        ExeVodCommand = blnRet
    Else
        strRetCmdSeq = -1
    End If
    
'    If strCMD <> "Z3" Then
'        strRetCmdSeqNo = -1
'    Else
'        strRetCmdSeq = strRetCmdSeqNo
'    End If
End Function
Private Function ExeOpenVod(ByRef rs As ADODB.Recordset, _
                             ByVal strCMD As String, _
                             Optional ByVal strCitemCode As String = Empty, _
                             Optional ByVal strCancelCitemCode As String = Empty, _
                             Optional ByVal strUpdTime As String = Empty, _
                             Optional ByVal strUpdEn As String = Empty, _
                             Optional ByRef strSTBUId As String = Empty, _
                             Optional ByVal strProcessingDate As String = Empty, _
                             Optional ByVal strSNO As String = Empty, _
                             Optional ByVal strMediaBillNo As String = Empty, _
                             Optional ByVal strReasonCode As String = Empty, _
                             Optional ByVal strReasonName As String = Empty, _
                             Optional ByRef rsSnd As ADODB.Recordset, _
                             Optional ByRef strResult As String, _
                             Optional ByRef strFaultReason As String, _
                             Optional ByRef strRetErrMsg As String, _
                             Optional ByVal strStartDate As String, _
                             Optional ByVal strStopDate As String, _
                             Optional ByVal strICCUID As String) As Boolean
  On Error GoTo ChkErr
    Dim rsVirtual As New ADODB.Recordset
    Dim rsOTT As New ADODB.Recordset
    ExeOpenVod = False
    blnShowMsg = False
    blnNoShowMsg = True
    Dim strChannelID As String
    Dim Notes As String
    Dim CitemCodes As String
    strChannelID = Empty
    OldVodAccountId = Empty
    Dim rsOld As New ADODB.Recordset
    If strCMD <> "Z4" And strCMD <> "Z5" Then
        Call GetSO001Data(rs)
    End If
    
Begin:
    Select Case UCase(strCMD)
        Case "A1"
            '#6874 記錄舊的VODAccountId,成功後要將舊值回填新值 By Kin 2014/09/30
            OldVodAccountId = rs("VodAccountId") & ""
            ExeOpenVod = OpenVod(rs, "A1", strUpdTime, , rs("CustId"), strCustName, strInstAddress, _
                                strZipCode, rs("FaciSNO") & "", "0000", rs("SmartCardNo") & "", _
                                GetMacAddress(rs("FaciSNO") & ""), rs("FaciSNO") & "", , , _
                                "", "", _
                                strUpdTime, strUpdEn, strSTBUId, strProcessingDate, strSNO, strMediaBillNo, _
                                strReasonCode, strReasonName, , , , , , , , GetTuner(rs("FaciSNo"), rs("SmartCardNo")), , _
                                rsVirtual, strResult, _
                                strRetErrMsg, strFaultReason)
        '#5713 增加填入ICCUID By Kin 2010/07/15
        Case "A2"
            ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , , , , rs("SmartCardNO") & "", _
                                    GetMacAddress(rs("FaciSNO") & ""), _
                                    rs("FaciSNo") & "", , rs("VODAccountId") & "", "", _
                                    "", strUpdTime, strUpdEn, strSTBUId, strProcessingDate, _
                                    strSNO, strMediaBillNo, strReasonCode, strReasonName, , , _
                                    GetICCUID(rs("SMARTCARDNO") & "", rs("COMPCODE")), _
                                    , , , , , , Nothing, strResult, strRetErrMsg, strFaultReason)
        Case "A3"
            ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , , _
                                     rs("FaciSNO") & "", , rs("SmartCardNo") & "", _
                                    , rs("FaciSNo") & "", , rs("VODAccountId") & "", _
                                    "", "", _
                                    strUpdTime, strUpdEn, , strProcessingDate, _
                                    strSNO, strMediaBillNo, strReasonCode, strReasonName, , , , , , , , , , _
                                     Nothing, strResult, strRetErrMsg, strFaultReason)
        Case "A4"
            ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , , rs("FaciSNO") & "", _
                                    , rs("SmartCardNo") & "", _
                                    , rs("FaciSNo") & "", , rs("VODAccountId") & "", _
                                    "", "", _
                                    strUpdTime, strUpdEn, , strProcessingDate, _
                                    strSNO, strMediaBillNo, strReasonCode, strReasonName, , , , , , , , , , _
                                    Nothing, strResult, strRetErrMsg, strFaultReason)
        '#5713 增加填入ICCUID By Kin 2010/07/15
        Case "A6"
            '因應NSTV專案判斷要送A2或A1或A6 By Kin 2013/07/30
            str7SNO = strSNO
            Dim OldFaciType As Integer
            Dim NewFaciType As Integer
            OldFaciType = GetDTVCAType(rs("FaciSNo") & "", rs("SmartCardNo") & "")
            NewFaciType = GetDTVCAType(rsSnd("FaciSNo") & "", rsSnd("SmartCardNo") & "")
            If OldFaciType <> NewFaciType Then
                '#6573 記錄舊的VodAccountId By Kin 2013/09/03
                OldVodAccountId = rs("VodAccountId")
                Set rsOld = rs.Clone
                Set rsOld.ActiveConnection = Nothing
                rsOld.AbsolutePosition = rs.AbsolutePosition
                Dim strProductID As String
                Dim strSO555BCitemCode As String
                Dim blnRet As Boolean
                strProductID = GetProductId(rs("VodAccountId"))
                strSO555BCitemCode = GetSO555BCitemCode(rs("VodAccountId"))
                '#6617 如果有SVOD頻道,要關掉之後再開 By Kin 2013/10/01
                If Len(strProductID) > 0 Then
                    blnRet = SetVodCmd(rs, "B12", strUpdTime, , rs("CustId"), , , , , , _
                                        rs("SmartCardNo") & "", _
                                        GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                        , , rs("VODAccountId") & "", _
                                         strProductID, , _
                                        strUpdTime, strUpdEn, , strProcessingDate, strSNO, _
                                        strMediaBillNo, strReasonCode, _
                                        strReasonName, , , , strSO555BCitemCode, strCancelCitemCode, Nothing, strResult, _
                                        strRetErrMsg, strFaultReason)
                Else
                    blnRet = True
                End If
                If blnRet Then
                    '#6981 增加預約時間間隔 By Kin 2015/01/20
                    If Len(strProcessingDate) > 0 Then
                        strProcessingDate = DateAdd("s", FPausetime, CDate(strProcessingDate))
                        strProcessingDate = Format(strProcessingDate, "yyyy/mm/dd hh:mm:ss")
                    End If
                    blnRet = OpenVod(rs, "A2", strUpdTime, , rs("CustId"), , , , , , rs("SmartCardNO") & "", _
                                  GetMacAddress(rs("FaciSNO") & ""), _
                                  rs("FaciSNo") & "", , rs("VODAccountId") & "", "", _
                                  "", strUpdTime, strUpdEn, strSTBUId, strProcessingDate, _
                                  strSNO, strMediaBillNo, strReasonCode, strReasonName, , , GetICCUID(rs("SMARTCARDNO") & "", rs("COMPCODE")), _
                                  , , , , , , _
                                   Nothing, strResult, strRetErrMsg, strFaultReason)
                    
                End If
                If blnRet Then
                     '#6981 增加預約時間間隔 By Kin 2015/01/20
                    If Len(strProcessingDate) > 0 Then
                        strProcessingDate = DateAdd("s", FPausetime, strProcessingDate)
                        strProcessingDate = Format(strProcessingDate, "yyyy/mm/dd hh:mm:ss")
                    End If
                    blnRet = OpenVod(rsSnd, "A1", strUpdTime, , rsSnd("CustId"), strCustName, strInstAddress, _
                               strZipCode, rsSnd("FaciSNO") & "", "0000", rsSnd("SmartCardNo") & "", _
                               GetMacAddress(rsSnd("FaciSNO") & ""), rsSnd("FaciSNO") & "", , , _
                               "", "", _
                               strUpdTime, strUpdEn, strSTBUId, strProcessingDate, strSNO, strMediaBillNo, _
                               strReasonCode, strReasonName, , , , , , , , GetTuner(rsSnd("FaciSNo"), rsSnd("SmartCardNo")), , _
                                rsVirtual, strResult, _
                               strRetErrMsg, strFaultReason)
                End If
                If blnRet Then
                    If Len(strProductID) > 0 Then
                        '#6981 增加預約時間間隔 By Kin 2015/01/20
                        If Len(strProcessingDate) > 0 Then
                            strProcessingDate = DateAdd("s", FPausetime, strProcessingDate)
                            strProcessingDate = Format(strProcessingDate, "yyyy/mm/dd hh:mm:ss")
                        End If
                        blnRet = SetVodCmd(rsSnd, "B11", strUpdTime, , rsSnd("CustId"), , , , , , _
                                        rsSnd("SmartCardNo") & "", _
                                        GetMacAddress(rsSnd("FaciSNo") & ""), rsSnd("FaciSNO") & "" _
                                        , , rsSnd("VODAccountId") & "", _
                                        strProductID, , _
                                        strUpdTime, strUpdEn, , strProcessingDate, strSNO, _
                                        strMediaBillNo, strReasonCode, _
                                        strReasonName, , , , strSO555BCitemCode, strCitemCode, Nothing, strResult, _
                                        strRetErrMsg, strFaultReason)
                    End If
                End If
                 '#6752 增加Smit命令(Nagra換NSTV設備時,要執行B13、B15、B17) By Kin 2014/05/13
                If OldFaciType = 0 And NewFaciType = 1 Then
                    Dim aSQL As String
                    Dim aCitemCode As String
                    Dim aCapacity As String
                    '執行B13
                    aSQL = GetRecordCitemCodeSQL(rsOld)
                    aCitemCode = GetMaxSizeCodeCitemCode(aSQL)
                    If Len(aCitemCode) > 0 Then
                         aCapacity = GetRecordCapacity(aCitemCode)
                         If aCapacity > 0 Then
                            '#6981 增加預約時間間隔 By Kin 2015/01/20
                            If Len(strProcessingDate) > 0 Then
                                strProcessingDate = DateAdd("s", FPausetime, strProcessingDate)
                                strProcessingDate = Format(strProcessingDate, "yyyy/mm/dd hh:mm:ss")
                            End If
                             blnRet = SetSmitVodCmd(rsSnd, "B13", strUpdTime, , rsSnd("CustId"), , , , , , _
                                    rsSnd("SmartCardNo") & "", _
                                    GetMacAddress(rsSnd("FaciSNo") & ""), rsSnd("FaciSNO") & "" _
                                    , , rsSnd("VODAccountId") & "", _
                                      , , _
                                    strUpdTime, , , strProcessingDate, strSNO, _
                                    str3MediaBillNo, strReasonCode, _
                                    strReasonName, , , , , , aCapacity, , strCitemCode, Nothing, strResult, _
                                    strRetErrMsg, strFaultReason)
                                                           
                         End If
                    End If
                    '執行B15
                    If blnRet Then
                        aSQL = GetSmitChargeCitemCodeSQL(rsOld, True, True, strCitemCode)
                        Notes = GetSmitProductId(aSQL, rsOld("CustId"), rsOld("FaciSNo"), rsOld("CompCode"), True, True)
                        If Len(Notes) > 0 Then
                             '#6981 增加預約時間間隔 By Kin 2015/01/20
                            If Len(strProcessingDate) > 0 Then
                                strProcessingDate = DateAdd("s", FPausetime, strProcessingDate)
                                strProcessingDate = Format(strProcessingDate, "yyyy/mm/dd hh:mm:ss")
                            End If
                           blnRet = SetSmitVodCmd(rsSnd, "B15", strUpdTime, , rsSnd("CustId"), , , , , , _
                                            rsSnd("SmartCardNo") & "", _
                                            GetMacAddress(rsSnd("FaciSNo") & ""), rsSnd("FaciSNO") & "" _
                                            , , rsSnd("VODAccountId") & "", _
                                              , , _
                                            strUpdTime, , , strProcessingDate, strSNO, _
                                            str3MediaBillNo, strReasonCode, _
                                            strReasonName, , , , , Notes, , , strCitemCode, Nothing, strResult, _
                                            strRetErrMsg, strFaultReason)
                        End If
                    End If
                    '執行B17
                    If blnRet Then
                        If IsCatchUp(rsOld) Then
                             '#6981 增加預約時間間隔 By Kin 2015/01/20
                            If Len(strProcessingDate) > 0 Then
                                strProcessingDate = DateAdd("s", FPausetime, strProcessingDate)
                                strProcessingDate = Format(strProcessingDate, "yyyy/mm/dd hh:mm:ss")
                            End If
                            blnRet = SetSmitVodCmd(rsSnd, "B17", strUpdTime, , rsSnd("CustId"), , , , , , _
                                rsSnd("SmartCardNo") & "", _
                                GetMacAddress(rsSnd("FaciSNo") & ""), rsSnd("FaciSNO") & "" _
                                , , rsSnd("VODAccountId") & "", _
                                  , , _
                                strUpdTime, , , strProcessingDate, strSNO, _
                                str3MediaBillNo, strReasonCode, _
                                strReasonName, , , , , , , , strCitemCode, Nothing, strResult, _
                                strRetErrMsg, strFaultReason)
                        End If
                    
                    End If
                End If
                ExeOpenVod = blnRet
            Else
                ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , , , , _
                                     rs("SmartCardNo") & "" & "," & rsSnd("SmartCardNo") & "", _
                                    GetMacAddress(rs("FaciSNO") & "") & "," & GetMacAddress(rsSnd("FaciSNO") & ""), _
                                     rsSnd("FaciSNO") & "", _
                                    , rs("VODAccountId") & "", _
                                    GetProductId(rs("VODAccountId")), "", _
                                    strUpdTime, strUpdEn, , strProcessingDate, _
                                    strSNO, strMediaBillNo, strReasonCode, _
                                    strReasonName, , , GetICCUID(rs("SMARTCARDNO") & "", rs("COMPCODE")), _
                                    rs("FaciSNo"), rs("SmartCardNo") & "", , , , , rsVirtual, strResult, _
                                    strRetErrMsg, strFaultReason)
            End If
        Case "A7"
            ExeOpenVod = SetSmitVodCmd(rs, "A7", strUpdTime, , rs("CustId"), , , , , , _
                                                rs("SmartCardNo") & "", _
                                                GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                                , , rs("VODAccountId") & "", _
                                                  , , strUpdTime, strUpdEn, , strProcessingDate, _
                                                  strSNO, _
                                                strMediaBillNo, strReasonCode, _
                                                strReasonName, , , , , , , GetTuner(rs("FaciSNo"), rs("SmartCardNo")), , Nothing, strResult, _
                                                strRetErrMsg, strFaultReason)
        Case "A8"
            ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , _
                                , , , rs("SmartCardNo") & "", _
                                GetMacAddress(rs("FaciSNO") & "") & "," & GetMacAddress(rsSnd("FaciSNO") & ""), _
                                rs("FaciSNo") & "", _
                                , rs("VODAccountId") & "", "", _
                                "", strUpdTime, strUpdEn, , strProcessingDate, _
                                strSNO, strMediaBillNo, strReasonCode, _
                                strReasonName, , , , , , , , , , Nothing, strResult, _
                                strRetErrMsg, strFaultReason)
                                
        Case "A9"
            ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , _
                                , , , rsSnd("SmartCardNo") & "", _
                                GetMacAddress(rs("FaciSNO") & ""), rs("FaciSNo") & "", _
                                , rs("VODAccountId") & "", "", "", strUpdTime, _
                                strUpdEn, , strProcessingDate, strSNO, strMediaBillNo, _
                                strReasonCode, strReasonName, , , , , , , , , , _
                                Nothing, strResult, strRetErrMsg, strFaultReason)
        Case "E1"
            ExeOpenVod = SetVodCmd(rs, strCMD, strUpdTime, , rs("CustId"), , , , rs("FaciSNO") & "", _
                                    "0000", rs("SmartCardNo") & "", , rs("FaciSNo") & "", _
                                    , rs("VODAccountId") & "", _
                                        , , strUpdTime, strUpdEn, , strProcessingDate, _
                                        strSNO, strMediaBillNo, strReasonCode, strReasonName, , , , , , _
                                        Nothing, strResult, strRetErrMsg, strFaultReason)
        Case "C4"
            ExeOpenVod = SetVodCmd(rs, strCMD, strUpdTime, , rs("CustId"), , , _
                            , , , rs("SmartCardNo") & "" _
                            , , rs("FaciSNo") & "", , rs("VODAccountId"), "", _
                            "", strUpdTime, strUpdEn, , strProcessingDate, _
                            strSNO, strMediaBillNo, strReasonCode, strReasonName, , , , , , _
                            Nothing, strResult, strRetErrMsg, strFaultReason)
        Case "Z2", "Z3"
            ExeOpenVod = SetVodCmd(rs, strCMD, strUpdTime, , rs("CustId") & "", , , , _
                                , , rs("SmartCardNo") & "", , rs("FaciSNo") & "" _
                                , , rs("VODAccountId") & "", , , strUpdTime, _
                                strUpdEn, , strProcessingDate, _
                                strSNO, strMediaBillNo, strReasonCode, _
                                strReasonName, strStartDate, strStopDate, , , , Nothing, _
                                strResult, strRetErrMsg, strFaultReason)
                                
        'Z4,Z5的MacAddress以明琪給我的為主,不然MacAddress跟FaciSNO都會變成同一個
        Case "Z4"
            ExeOpenVod = SetVodCmd(rs, strCMD, strUpdTime, , rs("CustId") & "", , , , , , _
                               rs("SmartCardNo") & "", rs("MacAddress") & "", _
                              rs("FaciSNo") & "", , rs("VODAccountId") & "", _
                            , , strUpdTime, strUpdEn, strSTBUId, strProcessingDate, _
                            strSNO, strMediaBillNo, strReasonCode, _
                            strReasonName, , , , , , rsVirtual, _
                           strResult, strRetErrMsg, strFaultReason)
        '#5603 增加傳入ICCUID By Kin 2010/03/29
        Case "Z5"
            ExeOpenVod = SetVodCmd(rs, strCMD, strUpdTime, , rs("CustId") & "", , , _
                        , , , rs("SmartCardNo") & "", _
                        rs("MacAddress") & "", rs("FaciSNO") & "", , _
                        rs("VODAccountId") & "", , , strUpdTime, _
                        strUpdEn, strSTBUId, strProcessingDate, _
                        strSNO, strMediaBillNo, strReasonCode, _
                        strReasonName, , , strICCUID, , , Nothing, _
                        strResult, strRetErrMsg, strFaultReason)
        Case "B11"
            '#5708 SO555A增加存入傳入的CitemCode By Kin 2010/07/05
             strChannelID = GetChannelID(strCitemCode, rs("CompCode"), rs("CustId"), 0, rs("VodAccountId") & "")
             If strChannelID = "" Then
                strRetErrMsg = "找不到任何ChannelID"
                ExeOpenVod = False
             Else
               ExeOpenVod = SetVodCmd(rs, "B11", strUpdTime, , rs("CustId"), , , , , , _
                                         rs("SmartCardNo") & "", GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                        , , rs("VODAccountId") & "", strChannelID, , _
                                        strUpdTime, strUpdEn, , strProcessingDate, strSNO, _
                                        strMediaBillNo, strReasonCode, _
                                        strReasonName, , , , strCitemCode, strCitemCode, Nothing, strResult, _
                                        strRetErrMsg, strFaultReason)
             End If
        Case "B12"
            '#5708 SO555A增加存入傳入的CitemCode By Kin 2010/07/05
             strChannelID = GetChannelID(strCancelCitemCode, rs("CompCode"), rs("CustId"), 1, rs("VodAccountId") & "")
             If strChannelID = "" Then
                strRetErrMsg = "找不到任何ChannelID"
                ExeOpenVod = False
             Else
                ExeOpenVod = SetVodCmd(rs, "B12", strUpdTime, , rs("CustId"), , , , , , _
                                        rs("SmartCardNo") & "", _
                                        GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                        , , rs("VODAccountId") & "", _
                                        strChannelID, , _
                                        strUpdTime, strUpdEn, , strProcessingDate, strSNO, _
                                        strMediaBillNo, strReasonCode, _
                                        strReasonName, , , , strCancelCitemCode, strCancelCitemCode, Nothing, strResult, _
                                        strRetErrMsg, strFaultReason)
             End If
         Case "B2B1"
            '#5708 SO555A增加存入傳入的CitemCode By Kin 2010/07/05
            Dim aRet As Boolean
            If strCancelCitemCode <> "" Then
                strChannelID = GetChannelID(strCancelCitemCode, rs("CompCode"), rs("CustId"), 1, rs("VodAccountId") & "")
                 If strChannelID = "" Then
                    'strRetErrMsg = "找不到任何ChannelID"
                    aRet = True
                    'ExeOpenVod = False
                 Else
                    aRet = SetVodCmd(rs, "B12", strUpdTime, , rs("CustId"), , , , , , _
                                            rs("SmartCardNo") & "", _
                                            GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                            , , rs("VODAccountId") & "", _
                                            strChannelID, , _
                                            strUpdTime, strUpdEn, , strProcessingDate, strSNO, _
                                            strMediaBillNo, strReasonCode, _
                                            strReasonName, , , , strCancelCitemCode, strCancelCitemCode, Nothing, strResult, _
                                            strRetErrMsg, strFaultReason)
                End If
            Else
                aRet = True
            End If
            If aRet Then
                If strCitemCode <> "" Then
                    strChannelID = GetChannelID(strCitemCode, rs("CompCode"), rs("CustId"), 0, rs("VodAccountId") & "")
                    If strChannelID = "" Then
                       'strRetErrMsg = "找不到任何ChannelID"
                       aRet = True
                       ExeOpenVod = True
                    Else
                      aRet = SetVodCmd(rs, "B11", strUpdTime, , rs("CustId"), , , , , , _
                                                rs("SmartCardNo") & "", GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                               , , rs("VODAccountId") & "", strChannelID, , _
                                               strUpdTime, strUpdEn, , strProcessingDate, strSNO, _
                                               strMediaBillNo, strReasonCode, _
                                               strReasonName, , , , strCitemCode, strCitemCode, Nothing, strResult, _
                                               strRetErrMsg, strFaultReason)
                    End If
                Else
                    aRet = True
                End If
            End If
            ExeOpenVod = aRet
        Case "B13"
            str7SNO = strSNO
            Dim DvrSize As String
            Dim maxSizeCitemCode As String
            If Len(strCitemCode) = 0 Then
                'DvrSize = GetChargeRecordCapacity(rs)
                ExeOpenVod = True
                Exit Function
            Else
                'DvrSize = GetRecordCapacity(strCitemCode)
                maxSizeCitemCode = GetMaxSizeCodeCitemCode(strCitemCode)
            End If
            If Len(maxSizeCitemCode) = 0 Then
                ExeOpenVod = True
                Exit Function
            End If
            DvrSize = GetRecordCapacity(maxSizeCitemCode)
            If DvrSize = 0 Then
                ExeOpenVod = True
                Exit Function
            End If
            ExeOpenVod = SetSmitVodCmd(rs, "B13", strUpdTime, , rs("CustId"), , , , , , _
                                                rs("SmartCardNo") & "", _
                                                GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                                , , rs("VODAccountId") & "", _
                                                  , , strUpdTime, strUpdEn, , strProcessingDate, _
                                                  strSNO, _
                                                strMediaBillNo, strReasonCode, _
                                                strReasonName, , , , , , DvrSize, , strCitemCode, Nothing, strResult, _
                                                strRetErrMsg, strFaultReason)
        
        Case "B14"
            str7SNO = strSNO
          
            If Len(strCancelCitemCode) = 0 Then
                ExeOpenVod = True
                Exit Function
            Else
                
                maxSizeCitemCode = GetMaxSizeCodeCitemCode(strCancelCitemCode)
                If Len(maxSizeCitemCode) = 0 Then
                    ExeOpenVod = True
                    Exit Function
                End If
                ExeOpenVod = SetSmitVodCmd(rs, "B14", strUpdTime, , rs("CustId"), , , , , , _
                                                rs("SmartCardNo") & "", _
                                                GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                                , , rs("VODAccountId") & "", _
                                                  , , strUpdTime, strUpdEn, , strProcessingDate, _
                                                  strSNO, _
                                                strMediaBillNo, strReasonCode, _
                                                strReasonName, , , , , , "0", , strCancelCitemCode, Nothing, strResult, _
                                                strRetErrMsg, strFaultReason)
            End If
                        
        Case "B15"
            str7SNO = strSNO
'            If Len(strCitemCode) = 0 Then
'                ExeOpenVod = True
'                Exit Function
'
'            Else
'                CitemCodes = strCitemCode
'            End If
            '#6803 收費項目除了傳進來的以外，還要加入該設備所有收費項目 By Kin 2014/05/30
            CitemCodes = GetSmitChargeCitemCodeSQL(rs, True, True, strCitemCode)
            If Len(CitemCodes) = 0 Then
                ExeOpenVod = True
                Exit Function
            End If
            Notes = GetSmitProductId(CitemCodes, rs("CustId"), rs("FaciSNo"), rs("CompCode"), True, False)
            If Len(Notes & "") = 0 Then
                ExeOpenVod = True
                Exit Function
            End If
            ExeOpenVod = SetSmitVodCmd(rs, "B15", strUpdTime, , rs("CustId"), , , , , , _
                                      rs("SmartCardNo") & "", _
                                      GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                      , , rs("VODAccountId") & "", _
                                        , , _
                                      strUpdTime, , , strProcessingDate, strSNO, _
                                      strMediaBillNo, strReasonCode, _
                                      strReasonName, , , , , Notes, , , strCitemCode, Nothing, strResult, _
                                          strRetErrMsg, strFaultReason)
        Case "B16"
            str7SNO = strSNO
            
            If Len(strCancelCitemCode) = 0 Then
                ExeOpenVod = True
                Exit Function
               
            Else
                CitemCodes = strCancelCitemCode
            End If
            Notes = GetSmitProductId(CitemCodes, rs("CustId"), rs("FaciSNo"), rs("CompCode"), True, True)
            If Len(Notes & "") = 0 Then
                ExeOpenVod = True
                Exit Function
            End If
            ExeOpenVod = SetSmitVodCmd(rs, "B16", strUpdTime, , rs("CustId"), , , , , , _
                                      rs("SmartCardNo") & "", _
                                      GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                      , , rs("VODAccountId") & "", _
                                        , , _
                                      strUpdTime, , , strProcessingDate, strSNO, _
                                      strMediaBillNo, strReasonCode, _
                                      strReasonName, , , , , Notes, , , strCancelCitemCode, Nothing, strResult, _
                                          strRetErrMsg, strFaultReason)
                                
        Case "B17"
            str7SNO = strSNO
'            CitemCodes = GetSmitChargeCitemCode(rs, True, False)
'            Notes = GetSmitProductId(CitemCodes)
'            If Len(Notes & "") = 0 Then
'                ExeOpenVod = True
'                Exit Function
'            End If
             
            If Len(strCitemCode) = 0 Then
                ExeOpenVod = True
                Exit Function
            Else
                If Not IsCatchUpCitemCode(strCitemCode) Then
                    ExeOpenVod = True
                    Exit Function
                End If
            End If
            
            ExeOpenVod = SetSmitVodCmd(rs, "B17", strUpdTime, , rs("CustId"), , , , , , _
                                  rs("SmartCardNo") & "", _
                                  GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                  , , rs("VODAccountId") & "", _
                                    , , _
                                  strUpdTime, , , strProcessingDate, strSNO, _
                                  strMediaBillNo, strReasonCode, _
                                  strReasonName, , , , , , , , strCitemCode, Nothing, strResult, _
                                      strRetErrMsg, strFaultReason)
        Case "B18"
            str7SNO = strSNO
            If Len(strCancelCitemCode) = 0 Then
                ExeOpenVod = True
                Exit Function
            Else
                If Not IsCatchUpCitemCode(strCancelCitemCode) Then
                    ExeOpenVod = True
                    Exit Function
                End If
            End If
           
            ExeOpenVod = SetSmitVodCmd(rs, "B18", strUpdTime, , rs("CustId"), , , , , , _
                                 rs("SmartCardNo") & "", _
                                 GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                 , , rs("VODAccountId") & "", _
                                   , , _
                                 strUpdTime, , , strProcessingDate, strSNO, _
                                 strMediaBillNo, strReasonCode, _
                                 strReasonName, , , , , , , , strCancelCitemCode, Nothing, strResult, _
                                     strRetErrMsg, strFaultReason)
            
        Case "C2"
            intOTTCompID = Val(GetRsValue("Select OTTCOMPID From " & GetOwner & "SO041") & "")
            If intOTTCompID <> 0 Then
                strOTTOwner = GetRsValue("Select TableOwner From " & GetOwner & "CD039 Where CodeNo = " & intOTTCompID) & "."
            End If
            If Len(strCitemCode & "") = 0 Then
                strRetErrMsg = "無傳入收費項目"
                ExeOpenVod = False
                Exit Function
            End If
             Dim strOTTUserID  As String
                    
            Dim aOTTSQL As String
            aOTTSQL = "Select A.OTTUserID,B.iOSOrderNo,B.FVOrderNo From " & strOTTOwner & "SO220 A," & strOTTOwner & "SO003 B  Where " & _
                                " A.OTTUserID In (Select OTTUserID From " & strOTTOwner & "SO222 Where  SOCompID = " & rs("COMPCODE") & _
                                " And FaciID = '" & rs("FaciID") & "' )  " & _
                                " AND A.CUSTID = B.CUSTID " & _
                                " AND B.FVORDERNO IS NOT NULL " & _
                                " AND B.CitemCode = " & strCitemCode
            
            GetRS rsOTT, aOTTSQL, gcnGi, adUseClient, adOpenKeyset
            If rsOTT.RecordCount = 0 Then
                strRetErrMsg = "找不到收費項目相關訂購資料"
                ExeOpenVod = False
                Exit Function
            End If
            If Len(rsOTT("OTTUserID") & "") = 0 Then
                strRetErrMsg = "找不到OTTUserID"
                ExeOpenVod = False
                Exit Function
            End If
            If Len(rsOTT("iOSOrderNo") & "") = 0 Then
                strRetErrMsg = "找不到iOS訂單單號訂購資料"
                ExeOpenVod = False
                Exit Function
            End If
            If Len(rsOTT("FVOrderNo") & "") = 0 Then
                strRetErrMsg = "找不到iFonsView訂單單號"
                ExeOpenVod = False
                Exit Function
            End If
            ExeOpenVod = QueryOTT(rs, "C2", strUpdTime, , rs("CustId"), , , , rs("FaciSNO") & "", _
                                         , rs("SmartCardNo") & "" _
                                        , , rs("FaciSNO") & "", , rs("VODAccountId") & "", _
                                        , , strUpdTime, , , strProcessingDate, strSNO, strMediaBillNo, _
                                        , , , , , , , _
                                        rsOTT("OTTUserID") & "", rsOTT("iOSOrderNo") & "", rsOTT("FVOrderNo") & "", Nothing, _
                                        strResult, strRetErrMsg, strFaultReason)
                                        
    End Select
    
    On Error Resume Next
    Call CloseRecordset(rsVirtual)
    Call CloseRecordset(rsOld)
    Call CloseRecordset(rsOTT)
    'strRetErrMsg = Empty
    'strFaultReason = Empty
    '#6140 特定錯誤原因要做返向指令 By Kin 2011/11/08
    If blnReturnCmd Then
        strRetErrMsg = Empty
        strFaultReason = Empty
        If UCase(strCMD) = "A1" Then
            strCMD = "A2"
            GoTo Begin
        End If
        If UCase(strCMD) = "A2" Then
            strCMD = "A1"
            GoTo Begin
        End If
    End If
    Exit Function
ChkErr:
  Call ErrSub("clsExeCommand", "ExeOpenVod")
End Function
Private Function GetSO001Data(ByRef rs As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr:
    Dim rsSO001 As New ADODB.Recordset
    GetSO001Data = False
    If Not GetRS(rsSO001, "Select CustName,InstAddrNo,InstAddress,TELH1,Tel1 From " & GetOwner & _
                    "So001 Where CustId = " & rs("CustId") & " And CompCode = " & rs("CompCode"), gcnGi) Then Exit Function
    If rs.EOF Then
        strCustName = ""
        lngInstAddrNo = 0
        strInstAddress = ""
        lngNodeNo = 0
        lngCircuitNo = 0
        strZipCode = ""
    Else
        strCustName = rsSO001("CustName") & ""
        strInstAddress = rsSO001("InstAddress") & ""
        strCustTel = IIf(rsSO001("TelH1") & "" = "", "", rsSO001("TelH1") & "-") & rsSO001("Tel1") & ""
        lngInstAddrNo = Val(rsSO001("InstAddrNo") & "")
        lngNodeNo = 0
        lngCircuitNo = 0
        strZipCode = GetRsValue("Select ZipCode From " & GetOwner & "SO014 Where AddrNo=" & rsSO001("InstAddrNo"), gcnGi) & ""
    End If
    GetSO001Data = True
    On Error Resume Next
    Call CloseRecordset(rsSO001)
    Exit Function
ChkErr:
  Call ErrSub("clsExeCommand", "GetSO001Data")
End Function
Private Sub Class_Terminate()
  On Error Resume Next
    objRecordWrite.Close
    Set objRecordWrite = Nothing
    Set objFileSystem = Nothing
    'CloseRecordset rs2
    
End Sub
