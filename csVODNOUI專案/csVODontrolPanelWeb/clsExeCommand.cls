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
                             Optional ByVal strICCUID As String = Empty) As Boolean
    On Error Resume Next
    blnShowMsg = False
    blnNoShowMsg = True
    ExeVodCommand = False
    blnExeCommand = blnExe '判斷是否執行命令再更新SO004
    
    Dim blnRet As Boolean
    Dim strArry As String
    
    strArry = Join(GaryGi, Chr(9))
    Set gcnGi = cn
    Call GetGlobal(strArry)
    If Not rsNew Is Nothing Then
        Set rs2 = rsNew
    Else
        Set rs2 = Nothing
    End If
    blnRecordProcedure = IsRecordProcedure  '判斷是否要記錄全部程序
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
    Call WriteRecordVodProcedure("3.完成執行命令程序並離開")
    If blnRet Then
        strRetCmdSeq = strRetCmdSeqNo
        ExeVodCommand = blnRet
    End If
    strRetCmdSeqNo = -1
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
  On Error GoTo chkErr
    Dim rsVirtual As New ADODB.Recordset
    ExeOpenVod = False
    blnShowMsg = False
    blnNoShowMsg = True
    Dim strChannelID As String
    strChannelID = Empty
    If strCMD <> "Z4" And strCMD <> "Z5" Then
        Call GetSO001Data(rs)
    End If
    
    Select Case UCase(strCMD)
        Case "A1"
            ExeOpenVod = OpenVod(rs, "A1", strUpdTime, , rs("CustId"), strCustName, strInstAddress, _
                                strZipCode, rs("FaciSNO") & "", "0000", rs("SmartCardNo") & "", _
                                GetMacAddress(rs("FaciSNO") & ""), rs("FaciSNO") & "", , , _
                                "", "", _
                                strUpdTime, strUpdEn, strSTBUId, strProcessingDate, strSNO, strMediaBillNo, _
                                strReasonCode, strReasonName, , , rsVirtual, strResult, _
                                strRetErrMsg, strFaultReason)
        '#5713 增加填入ICCUID By Kin 2010/07/15
        Case "A2"
            ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , , , , rs("SmartCardNO") & "", _
                                    GetMacAddress(rs("FaciSNO") & ""), _
                                    rs("FaciSNo") & "", , rs("VODAccountId") & "", "", _
                                    "", strUpdTime, strUpdEn, strSTBUId, strProcessingDate, _
                                    strSNO, strMediaBillNo, strReasonCode, strReasonName, , , _
                                    Nothing, strResult, strRetErrMsg, strFaultReason, _
                                    GetICCUID(rs("SMARTCARDNO") & "", rs("COMPCODE")))
        Case "A3"
            ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , , _
                                     rs("FaciSNO") & "", , rs("SmartCardNo") & "", _
                                    , rs("FaciSNo") & "", , rs("VODAccountId") & "", _
                                    "", "", _
                                    strUpdTime, strUpdEn, , strProcessingDate, _
                                    strSNO, strMediaBillNo, strReasonCode, strReasonName, , , _
                                     Nothing, strResult, strRetErrMsg, strFaultReason)
        Case "A4"
            ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , , rs("FaciSNO") & "", _
                                    , rs("SmartCardNo") & "", _
                                    , rs("FaciSNo") & "", , rs("VODAccountId") & "", _
                                    "", "", _
                                    strUpdTime, strUpdEn, , strProcessingDate, _
                                    strSNO, strMediaBillNo, strReasonCode, strReasonName, , , _
                                    Nothing, strResult, strRetErrMsg, strFaultReason)
        '#5713 增加填入ICCUID By Kin 2010/07/15
        Case "A6"
            ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , , , , _
                                 rs("SmartCardNo") & "" & "," & rsSnd("SmartCardNo") & "", _
                                GetMacAddress(rs("FaciSNO") & "") & "," & GetMacAddress(rsSnd("FaciSNO") & ""), _
                                 rsSnd("FaciSNO") & "", _
                                , rs("VODAccountId") & "", _
                                GetProductID(rs("VODAccountId")), "", _
                                strUpdTime, strUpdEn, , strProcessingDate, _
                                strSNO, strMediaBillNo, strReasonCode, _
                                strReasonName, , , rsVirtual, strResult, _
                                strRetErrMsg, strFaultReason, GetICCUID(rs("SMARTCARDNO") & "", rs("COMPCODE")))
        Case "A7"
            ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , , _
                                , , rs("SmartCardNo") & "", _
                                GetMacAddress(rs("FaciSNO") & ""), rs("FaciSNo") & "" _
                                , , rs("VODAccountId") & "", "", "", _
                                strUpdTime, strUpdEn, , strProcessingDate, _
                                strSNO, strMediaBillNo, strReasonCode, _
                                strReasonName, , , Nothing, strResult, _
                                strRetErrMsg, strFaultReason)
        Case "A8"
            ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , _
                                , , , rs("SmartCardNo") & "", _
                                GetMacAddress(rs("FaciSNO") & "") & "," & GetMacAddress(rsSnd("FaciSNO") & ""), _
                                rs("FaciSNo") & "", _
                                , rs("VODAccountId") & "", "", _
                                "", strUpdTime, strUpdEn, , strProcessingDate, _
                                strSNO, strMediaBillNo, strReasonCode, _
                                strReasonName, , , Nothing, strResult, _
                                strRetErrMsg, strFaultReason)
                                
        Case "A9"
            ExeOpenVod = OpenVod(rs, strCMD, strUpdTime, , rs("CustId"), , , _
                                , , , rsSnd("SmartCardNo") & "", _
                                GetMacAddress(rs("FaciSNO") & ""), rs("FaciSNo") & "", _
                                , rs("VODAccountId") & "", "", "", strUpdTime, _
                                strUpdEn, , strProcessingDate, strSNO, strMediaBillNo, _
                                strReasonCode, strReasonName, , , _
                                Nothing, strResult, strRetErrMsg, strFaultReason)
        Case "E1"
            ExeOpenVod = SetVodCmd(rs, strCMD, strUpdTime, , rs("CustId"), , , , rs("FaciSNO") & "", _
                                    "0000", rs("SmartCardNo") & "", , rs("FaciSNo") & "", _
                                    , rs("VODAccountId") & "", _
                                        , , strUpdTime, strUpdEn, , strProcessingDate, _
                                        strSNO, strMediaBillNo, strReasonCode, strReasonName, , , _
                                        Nothing, strResult, strRetErrMsg, strFaultReason)
        Case "C4"
            ExeOpenVod = SetVodCmd(rs, strCMD, strUpdTime, , rs("CustId"), , , _
                            , , , rs("SmartCardNo") & "" _
                            , , rs("FaciSNo") & "", , rs("VODAccountId"), "", _
                            "", strUpdTime, strUpdEn, , strProcessingDate, _
                            strSNO, strMediaBillNo, strReasonCode, strReasonName, , , _
                            Nothing, strResult, strRetErrMsg, strFaultReason)
        Case "Z2", "Z3"
            ExeOpenVod = SetVodCmd(rs, strCMD, strUpdTime, , rs("CustId") & "", , , , _
                                , , rs("SmartCardNo") & "", , rs("FaciSNo") & "" _
                                , , rs("VODAccountId") & "", , , strUpdTime, _
                                strUpdEn, , strProcessingDate, _
                                strSNO, strMediaBillNo, strReasonCode, _
                                strReasonName, strStartDate, strStopDate, Nothing, _
                                strResult, strRetErrMsg, strFaultReason)
                                
        'Z4,Z5的MacAddress以明琪給我的為主,不然MacAddress跟FaciSNO都會變成同一個
        Case "Z4"
            ExeOpenVod = SetVodCmd(rs, strCMD, strUpdTime, , rs("CustId") & "", , , , , , _
                               rs("SmartCardNo") & "", rs("MacAddress") & "", _
                              rs("FaciSNo") & "", , rs("VODAccountId") & "", _
                            , , strUpdTime, strUpdEn, strSTBUId, strProcessingDate, _
                            strSNO, strMediaBillNo, strReasonCode, _
                            strReasonName, , , rsVirtual, _
                           strResult, strRetErrMsg, strFaultReason)
        '#5603 增加傳入ICCUID By Kin 2010/03/29
        Case "Z5"
            ExeOpenVod = SetVodCmd(rs, strCMD, strUpdTime, , rs("CustId") & "", , , _
                        , , , rs("SmartCardNo") & "", _
                        rs("MacAddress") & "", rs("FaciSNO") & "", , _
                        rs("VODAccountId") & "", , , strUpdTime, _
                        strUpdEn, strSTBUId, strProcessingDate, _
                        strSNO, strMediaBillNo, strReasonCode, _
                        strReasonName, , , Nothing, _
                        strResult, strRetErrMsg, strFaultReason, strICCUID)
        Case "B11"
            '#5708 SO555A增加存入傳入的CitemCode By Kin 2010/07/05
             strChannelID = GetChannelID(strCitemCode, rs("CompCode"), rs("CustId"), 0, rs("VodAccountId"))
             If strChannelID = "" Then
                strRetErrMsg = "找不到任何ChannelID"
                ExeOpenVod = False
             Else
               ExeOpenVod = SetVodCmd(rs, "B11", strUpdTime, , rs("CustId"), , , , , , _
                                         rs("SmartCardNo") & "", GetMacAddress(rs("FaciSNo") & ""), rs("FaciSNO") & "" _
                                        , , rs("VODAccountId") & "", strChannelID, , _
                                        strUpdTime, strUpdEn, , strProcessingDate, strSNO, _
                                        strMediaBillNo, strReasonCode, _
                                        strReasonName, , , Nothing, strResult, _
                                        strRetErrMsg, strFaultReason, , strCitemCode)
             End If
        Case "B12"
            '#5708 SO555A增加存入傳入的CitemCode By Kin 2010/07/05
             strChannelID = GetChannelID(strCancelCitemCode, rs("CompCode"), rs("CustId"), 1, rs("VodAccountId"))
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
                                        strReasonName, , , Nothing, strResult, _
                                        strRetErrMsg, strFaultReason, , strCancelCitemCode)
             End If
         Case "B2B1"
            '#5708 SO555A增加存入傳入的CitemCode By Kin 2010/07/05
            Dim aRet As Boolean
            If strCancelCitemCode <> "" Then
                strChannelID = GetChannelID(strCancelCitemCode, rs("CompCode"), rs("CustId"), 1, rs("VodAccountId"))
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
                                            strReasonName, , , Nothing, strResult, _
                                            strRetErrMsg, strFaultReason, , strCancelCitemCode)
                End If
            Else
                aRet = True
            End If
            If aRet Then
                If strCitemCode <> "" Then
                    strChannelID = GetChannelID(strCitemCode, rs("CompCode"), rs("CustId"), 0, rs("VodAccountId"))
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
                                               strReasonName, , , Nothing, strResult, _
                                               strRetErrMsg, strFaultReason, , strCitemCode)
                    End If
                Else
                    aRet = True
                End If
            End If
            ExeOpenVod = aRet
    End Select
    
    On Error Resume Next
    Call CloseRecordset(rsVirtual)
    Exit Function
chkErr:
  Call ErrSub("clsExeCommand", "ExeOpenVod")
End Function
Private Function GetSO001Data(ByRef rs As ADODB.Recordset) As Boolean
  On Error GoTo chkErr:
    Dim rsSO001 As New ADODB.Recordset
    GetSO001Data = False
    If Not GetRS(rsSO001, "Select CustName,InstAddrNo,InstAddress,TELH1,Tel1 From " & GetOwner & _
                    "So001 Where CustId = " & rs("CustId") & " And CompCode = " & rs("CompCode"), gcnGi, , , , , , , True) Then Exit Function
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
chkErr:
  Call ErrSub("clsExeCommand", "GetSO001Data")
End Function



Private Sub Class_Terminate()
  On Error Resume Next
    objRecordWrite.Close
    Set objRecordWrite = Nothing
    Set objFileSystem = Nothing
    'CloseRecordset rs2
    
End Sub
