Attribute VB_Name = "mduControlPanelWeb"
Option Explicit

Private Const blnDebug = False

Public gCompFilterStr As String


Public gDefaultComp As String
Public objActForm As Form
Public gLogInID As String
Public gstrUserPremission As String
Public lngSeqNo As Long
Public strOwner As String
'Dim lngCMtimeout As Long
Public strCmdSeqNo As String
Public blnExeCommand As Boolean
Public strCMowner As String
Dim strErrorMessage As String
Public intCompCode As Integer
Public blnShowMsg As Boolean
Public IsA6Command As Boolean
Public IsRestCommand As Boolean
Public OldVodAccountId As String
Public strFaciSNo As String
Public strFaciCode As String
Public strResvTime As String
Public blnSendKey As Boolean
Public strReason As String
Public blnShowFaci As Boolean
Public str7SNO As String
Public str3MediaBillNo As String
Public strSeqNo As String
Public strReasonCode As String
Public strReasonName As String
Public strRetCmdSeqNo As String
Public lngCDRTimeOut As Long
Public blnRecordProcedure As Boolean
Public blnReturnCmd As Boolean
Public blnHasRunReturn As Boolean
Public intRunRetunCount As Integer
Public Const strA1ErrCode = "120116"
Public Const strA2ErrCode = "120120"
Public rs2 As New ADODB.Recordset
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Function GetModeTypeValue(ByVal strValue As String) As String
    On Error Resume Next
    Dim strReturn As String
        
        Select Case strValue
            Case "A1"
                strReturn = "開機"
            Case "A2"
                strReturn = "關機"
            Case "A3"
                strReturn = "停機"
            Case "A4"
                strReturn = "重設(Reset)"
            
            Case "B1"
                strReturn = "換CM"
            Case "B2"
                strReturn = "變更CM速率"
            Case "B3"
                strReturn = "Ping CM"
            Case "B4"
                strReturn = "釋放動態IP"
            Case "B5"
                strReturn = "ACS帳號密碼相關功能"
            
            Case "C1"
                strReturn = "CM 狀態查詢"
            Case "C2"
                strReturn = "查詢 HFC 服務類別"
            Case "C3"
                strReturn = "清除EMM"
            Case "C4"
                strReturn = "CM 連線紀錄查詢"
            Case "C5"
                strReturn = "PCIP連線記錄查詢"
        End Select
        
        GetModeTypeValue = strReturn
End Function
Public Function GetChannelCitemCode(ByVal aChannelID As String, ByVal aCitemCode As String) As String
  On Error GoTo ChkErr
    Dim rsCD019A As New ADODB.Recordset
    Dim aSQL As String
    If aCitemCode = "" Then
        GetChannelCitemCode = ""
        Exit Function
    End If
    aSQL = "SELECT A.CITEMCODE FROM " & GetOwner & "CD019A A," & GetOwner & "CD024 B " & _
            " WHERE B.CHANNELID ='" & aChannelID & "' AND A.CODENO = B.CODENO " & _
            " AND B.VODTYPE=2 AND B.STOPFLAG <> 1 " & _
            " AND A.CITEMCODE IN(" & Replace(aCitemCode, "'", "") & ")"
    If Not GetRS(rsCD019A, aSQL, gcnGi, adUseClient, adOpenKeyset) Then Exit Function
    If rsCD019A.EOF Then
        GetChannelCitemCode = ""
    Else
        GetChannelCitemCode = rsCD019A("CITEMCODE") & ""
    End If
    On Error Resume Next
    Call CloseRecordset(rsCD019A)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "GetChannelCitemCode")
End Function

'#5692 取得SVOD頻道(aStatusType=0 訂購SVOD) By Kin 2010/06/15
Public Function GetChannelID(ByVal aCitemCode As String, ByVal aCompCode As String, _
                            Optional ByVal aCustId As String = Empty, _
                            Optional ByVal aStatusType As Integer = 0, _
                            Optional ByVal aVodAccountId As String) As String
  On Error GoTo ChkErr
    Dim rsCD024 As New ADODB.Recordset
    Dim aSQL As String
    Dim aRet As String
    aSQL = Empty
    aRet = Empty
    If aCitemCode = "" Then
        GetChannelID = ""
        Exit Function
    End If
    aCitemCode = Replace(aCitemCode, "'", "")
    aSQL = "SELECT DISTINCT CD024.ChannelID FROM " & GetOwner & "CD024 " & _
        " WHERE CD024.CodeNo IN( SELECT CODENO FROM " & GetOwner & "CD019A " & _
        " WHERE CD019A.CITEMCODE IN(" & aCitemCode & ")) " & _
        " AND CD024.VODType=2 AND CD024.STOPFLAG <> 1 AND CD024.COMPCODE=" & aCompCode
    '如果沒VodAccountId就不要串下面的語法了，不然會錯 By Kin 2014/06/05 For Corey
    If Len(aVodAccountId) > 0 Then
        If aStatusType = 0 Then '0代表要開通SVOD 1代表要關SVOD
            aSQL = aSQL & " AND CD024.CODENO NOT IN(SELECT ProductID FROM " & GetOwner & "SO555B " & _
                    " WHERE NVL(STATUS,0)=1 AND VODACCOUNTID = " & aVodAccountId & ") "
        Else
            aSQL = aSQL & " AND CD024.CODENO NOT IN(SELECT ProductID FROM " & GetOwner & "SO555B " & _
                    " WHERE NVL(STATUS,0)=0 AND VODACCOUNTID = " & aVodAccountId & ") "
        End If
    End If
    If Not GetRS(rsCD024, aSQL, gcnGi, adUseClient, adOpenKeyset, , , , , False) Then Exit Function
    Do While Not rsCD024.EOF
        If aRet = "" Then
            aRet = rsCD024("ChannelID") & ""
        Else
            aRet = aRet & "," & rsCD024("ChannelID")
        End If
        rsCD024.MoveNext
    Loop
    GetChannelID = aRet
    On Error Resume Next
    Call CloseRecordset(rsCD024)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "GetChannelID")
End Function
Public Function Get_EMTA_Mac(strFaciSNo As String) As String
  On Error GoTo ChkErr
    If strFaciSNo <> Empty Then
        Get_EMTA_Mac = GetRsValue("SELECT MTAMAC FROM " & GetOwner & "SO306" & _
                                                        " WHERE HFCMAC='" & strFaciSNo & "'") & ""
       
    End If
  Exit Function
  
ChkErr:
    ErrSub FormName, "Get_EMTA_Mac"
End Function
Public Function Get_EMTA_Mode(strFaciSNo As String) As String
  On Error Resume Next
    Get_EMTA_Mode = GetRsValue("SELECT Nvl(ModeID,0) FROM " & GetOwner & "SO306" & _
                                                        " WHERE HFCMAC='" & strFaciSNo & "'") & ""
End Function
Public Function GetMacAddress(ByVal aFaciSNo As String) As String
  On Error GoTo ChkErr
    Dim rsMac As New ADODB.Recordset
    Dim strMac As String
    Dim strTmp As String
    Dim i As Integer
    If aFaciSNo = "" Then
        strMac = ""
    Else
        If Not GetRS(rsMac, "Select EthernetMAC From " & GetOwner & "SOAC0201A Where " & _
                        "FaciSNo='" & aFaciSNo & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then Exit Function
        If rsMac.EOF Then
            strMac = ""
        Else
            If rsMac(0) & "" = "" Then
                strMac = ""
            Else
                strMac = rsMac(0).Value
                If InStr(1, strMac, ":") = 0 Then
                    For i = 2 To Len(strMac) Step 2
                        strTmp = strTmp & Mid(strMac, i - 1, 2) & ":"
                    Next
                    strMac = Mid(strTmp, 1, Len(strTmp) - 1)
                End If
            End If
        End If
    End If
    GetMacAddress = strMac
  On Error Resume Next
  Call CloseRecordset(rsMac)
  Exit Function
ChkErr:
  ErrSub FormName, "GetMacAddress"
End Function
'#6752 增加7個Smit Vod命令 By Kin 2014/04/08
Public Function SetSmitVodCmd(ByRef rsData As ADODB.Recordset, ByVal strCMD As String, _
                        Optional ByVal strExchTime As String = "", Optional ByVal strServiceProviderID As String = "", _
                        Optional ByVal strAccountNum As String, Optional ByVal strAccoutName As String, _
                        Optional ByVal strAddress As String = "", Optional ByVal strPostCode As String = "", _
                        Optional ByVal strLoginID As String = "", Optional ByVal strPin As String = "", _
                        Optional ByVal strSmartCardID As String = "", _
                        Optional ByVal strMacAddress As String = "", Optional ByVal strSerialNumber As String = "", _
                        Optional ByVal strCASNo As String = "", Optional ByVal strAccountUID As String, _
                        Optional ByVal strProductID As String = "", Optional ByVal strProductName As String, _
                        Optional ByVal strUpdTime As String = "", Optional ByVal strUpdEn As String = "", _
                        Optional ByRef strSTBUId As String = "", _
                        Optional ByVal strProcessingDate As String = "", Optional ByVal strSNO As String, _
                        Optional ByVal strMediaBillNo As String = "", Optional ByVal strReasonCode As String = "", _
                        Optional ByVal strReasonName As String = "", Optional ByVal strBillStartDate As String = "", _
                        Optional ByVal strBillEndDate As String = "", _
                        Optional ByVal strICCUID As String, Optional ByVal strCitemStr As String = "", _
                        Optional ByVal strNotes As String = "", Optional strDVRQuota As String = "", Optional strTunerCount As String = "", _
                         Optional ByVal strCitemCodes As String = "", _
                        Optional ByRef rsResult As Object = Nothing, Optional strResult As String, _
                        Optional ByRef strRetErrMsg As String, Optional ByRef strFaultReason As String) As Boolean
    On Error GoTo ChkErr
       
        Dim varBookMark As Variant
        Dim strCmdName As String
        Dim strUpdDate As String
        Dim strUpdName As String
        Dim objVodCommand As New clsVodcommand
        Dim intPVRCMDType As Byte
        intPVRCMDType = 0
        If strExchTime = "" Then
            strUpdDate = RightNow
        Else
            strUpdDate = strExchTime
        End If
        
        If strUpdTime = "" Then
            strUpdTime = strUpdDate
        End If
        If strUpdEn = "" Then
            strUpdName = GaryGi(1)
        Else
            strUpdName = strUpdEn
        End If
        intCompCode = rsData("CompCode")
        Call WriteRecordVodProcedure("2.2 進入新增控制台資料 --> " & strCMD)
        varBookMark = rsData.AbsolutePosition
        
        Select Case UCase(strCMD)
            Case "A7"
                strCmdName = "設定機上盒Tuner數量"
            Case "B13"
                strCmdName = "設定錄影容量"
            Case "B14"
                strCmdName = "取消錄影容量"
            Case "B15"
                strCmdName = "機上盒開通套餐Prod ID"
            Case "B16"
                strCmdName = "機上盒關閉套餐Prod ID"
            Case "B17"
                strCmdName = "開啟CatchUp"
            Case "B18"
                strCmdName = "關閉CatchUp"
        End Select
        '反向即時命令如果VodAccountId是空值,則不要送命令,直接回成功 For Jacky By Kin 2014/0716
        If (IsCancelCmd) And (Len(strProcessingDate & "") = 0) And (Len(rsData("VodAccountId") & "") = 0) Then
            strResult = "1"
            SetSmitVodCmd = True
            Call WriteRecordVodProcedure("2.2 反向即時而且VodAccountId無值，不送命令! --> " & strCMD)
            GoTo 88
        End If
        '增加判斷Jacky jacky: CD043.PVRCMDType = 0 才送命令 By Kin 2015/05/07
        intPVRCMDType = qryPVRCMDType(rsData("ModelCode") & "")
        If intPVRCMDType = 0 Then
            If blnExeCommand Then
                Call WriteRecordVodProcedure("2.2 進入新增控制台資料 --> " & strCMD)
                If Not objVodCommand.Vod_Command(gcnGi.ConnectionString, intCompCode & "", GetOwner, strCMD _
                                        , , , , , strUpdDate, , , strServiceProviderID, strAccountNum, _
                                        strAccoutName, , strAddress, strPostCode, _
                                        strAccountUID, strLoginID, _
                                        strPin, , , strSmartCardID, strMacAddress, _
                                        strSerialNumber, , , , strCASNo _
                                        , strProductID, strProductName, strUpdTime, strUpdName, _
                                        strSTBUId, strProcessingDate, _
                                        strSNO, strMediaBillNo, strReasonCode, strReasonName, GaryGi(0), _
                                        str(intCompCode), strBillStartDate, strBillEndDate, _
                                        strICCUID, strCitemStr, , , strNotes, strDVRQuota, strTunerCount, strCitemCodes, _
                                        , , , GetPrbCtl, rsResult, strResult, strRetErrMsg, _
                                        strFaultReason) Then GoTo 88
    
                Call WriteRecordVodProcedure("2.3 完成新增控制台資料 --> " & strCMD)
            End If
        Else
            blnShowMsg = False
        End If
        If strProcessingDate <> Empty And blnExeCommand Then
            strResult = "1"
            SetSmitVodCmd = True
            Call WriteRecordVodProcedure("2.4 預約資料 --> " & strCMD)
            GoTo 88
        End If
        strResult = "1"
        SetSmitVodCmd = True
        Call WriteRecordVodProcedure("2.5 完成更新資料 --> " & strCMD)
88:
        If blnShowMsg Then
            If IsA6Command Then
            
            Else
                ShowMsg "(2) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
            End If
        End If
        On Error Resume Next
        'Call CloseRecordset(rs555)
        Exit Function
     Exit Function
ChkErr:
  Call ErrSub(FormName, "SetSmitVodCmd")
End Function
Public Function QueryOTT(ByRef rsData As ADODB.Recordset, ByVal strCMD As String, _
                        Optional ByVal strExchTime As String = "", Optional ByVal strServiceProviderID As String = "", _
                        Optional ByVal strAccountNum As String, Optional ByVal strAccoutName As String, _
                        Optional ByVal strAddress As String = "", Optional ByVal strPostCode As String = "", _
                        Optional ByVal strLoginID As String = "", Optional ByVal strPin As String = "", _
                        Optional ByVal strSmartCardID As String = "", _
                        Optional ByVal strMacAddress As String = "", Optional ByVal strSerialNumber As String = "", _
                        Optional ByVal strCASNo As String = "", Optional ByVal strAccountUID As String, _
                        Optional ByVal strProductID As String = "", Optional ByVal strProductName As String, _
                        Optional ByVal strUpdTime As String = "", Optional ByVal strUpdEn As String = "", _
                        Optional ByRef strSTBUId As String = "", _
                        Optional ByVal strProcessingDate As String = "", Optional ByVal strSNO As String, _
                        Optional ByVal strMediaBillNo As String = "", Optional ByVal strReasonCode As String = "", _
                        Optional ByVal strReasonName As String = "", Optional ByVal strBillStartDate As String = "", _
                        Optional ByVal strBillEndDate As String = "", _
                        Optional ByVal strICCUID As String, Optional ByVal strCitemStr As String = "", _
                        Optional ByVal strCitemCodes As String = "", _
                        Optional ByVal strOTTUserID As String = "", _
                        Optional ByVal striOSOrderNo As String = "", _
                        Optional ByVal strFVOrderNo As String = "", _
                        Optional ByRef rsResult As Object = Nothing, Optional strResult As String, _
                        Optional ByRef strRetErrMsg As String, Optional ByRef strFaultReason As String) As Boolean
                        
    Dim strNowDate As String
    Dim varBookMark As Variant
    Dim strCmdName As String
    Dim strUpdDate As String
    Dim strUpdName As String
    Dim objVodCommand As New clsVodcommand
    'Dim rs555B As New ADODB.Recordset
    Dim rs555 As New ADODB.Recordset
    'Dim rs005B As New ADODB.Recordset
    Dim i As Integer
    Dim aCitemCode As String
    Dim aServiceType As String
    Dim strUPDSOAC0201B As String
    If strExchTime = "" Then
        strUpdDate = RightNow
    Else
        strUpdDate = strExchTime
    End If
    
    If strUpdTime = "" Then
        strUpdTime = strUpdDate
    End If
    If strUpdEn = "" Then
        strUpdName = GaryGi(1)
    Else
        strUpdName = strUpdEn
    End If
    intCompCode = rsData("CompCode")
    Call WriteRecordVodProcedure("2.2 進入新增控制台資料 --> " & strCMD)
    varBookMark = rsData.AbsolutePosition
    Select Case UCase(strCMD)
        Case "C2"
            strCmdName = "IOS刷卡狀態查詢"
    End Select
    If (IsCancelCmd) And (Len(strProcessingDate & "") = 0) And (Len(rsData("VodAccountId") & "") = 0) Then
        strResult = "1"
        QueryOTT = True
        Call WriteRecordVodProcedure("2.2 反向即時而且VodAccountId無值，不送命令! --> " & strCMD)
        GoTo 88
    End If
    If blnExeCommand Then
        Call WriteRecordVodProcedure("2.2 進入新增控制台資料 --> " & strCMD)
        If Not objVodCommand.Vod_Command(gcnGi.ConnectionString, intCompCode & "", GetOwner, strCMD _
                                , , , , , strUpdDate, , , strServiceProviderID, strAccountNum, _
                                strAccoutName, , strAddress, strPostCode, _
                                strAccountUID, strLoginID, _
                                strPin, , , strSmartCardID, strMacAddress, _
                                strSerialNumber, , , , strCASNo _
                                , strProductID, strProductName, strUpdTime, strUpdName, _
                                strSTBUId, strProcessingDate, _
                                strSNO, strMediaBillNo, strReasonCode, strReasonName, GaryGi(0), _
                                str(intCompCode), strBillStartDate, strBillEndDate, _
                                strICCUID, strCitemStr, , , , , , strCitemCodes, _
                                strOTTUserID, striOSOrderNo, strFVOrderNo, GetPrbCtl, rsResult, strResult, strRetErrMsg, _
                                strFaultReason) Then GoTo 88
            
        Call WriteRecordVodProcedure("2.3 完成新增控制台資料 --> " & strCMD)
    End If
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        QueryOTT = True
        Call WriteRecordVodProcedure("2.4 預約資料 --> " & strCMD)
        GoTo 88
    End If
    rsData.AbsolutePosition = varBookMark
    strResult = "1"
    QueryOTT = True
    If blnShowMsg Then
         If Len(strRetErrMsg & "") > 0 Then
            MsgBox strRetErrMsg, vbInformation, "訊息"
        End If
    End If
   
88:
    If blnShowMsg Then
        If IsA6Command Or IsRestCommand Then
'            If UCase(strCMD) <> "B11" Then
'                ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
'            End If
        Else
            ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
        End If
    End If
    
    On Error Resume Next
    'Call CloseRecordset(rs005)
    'Call CloseRecordset(rs555)
    'Call CloseRecordset(rs005B)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "QueryOTT")
End Function


Public Function SetVodCmd(ByRef rsData As ADODB.Recordset, ByVal strCMD As String, _
                        Optional ByVal strExchTime As String = "", Optional ByVal strServiceProviderID As String = "", _
                        Optional ByVal strAccountNum As String, Optional ByVal strAccoutName As String, _
                        Optional ByVal strAddress As String = "", Optional ByVal strPostCode As String = "", _
                        Optional ByVal strLoginID As String = "", Optional ByVal strPin As String = "", _
                        Optional ByVal strSmartCardID As String = "", _
                        Optional ByVal strMacAddress As String = "", Optional ByVal strSerialNumber As String = "", _
                        Optional ByVal strCASNo As String = "", Optional ByVal strAccountUID As String, _
                        Optional ByVal strProductID As String = "", Optional ByVal strProductName As String, _
                        Optional ByVal strUpdTime As String = "", Optional ByVal strUpdEn As String = "", _
                        Optional ByRef strSTBUId As String = "", _
                        Optional ByVal strProcessingDate As String = "", Optional ByVal strSNO As String, _
                        Optional ByVal strMediaBillNo As String = "", Optional ByVal strReasonCode As String = "", _
                        Optional ByVal strReasonName As String = "", Optional ByVal strBillStartDate As String = "", _
                        Optional ByVal strBillEndDate As String = "", _
                        Optional ByVal strICCUID As String, Optional ByVal strCitemStr As String = "", _
                        Optional ByVal strCitemCodes As String = "", _
                        Optional ByRef rsResult As Object = Nothing, Optional strResult As String, _
                        Optional ByRef strRetErrMsg As String, Optional ByRef strFaultReason As String) As Boolean
  On Error GoTo ChkErr
    Dim strNowDate As String
    Dim varBookMark As Variant
    Dim strCmdName As String
    Dim strUpdDate As String
    Dim strUpdName As String
    Dim objVodCommand As New clsVodcommand
    'Dim rs555B As New ADODB.Recordset
    Dim rs555 As New ADODB.Recordset
    'Dim rs005B As New ADODB.Recordset
    Dim i As Integer
    Dim aCitemCode As String
    Dim aServiceType As String
    Dim strUPDSOAC0201B As String
    If strExchTime = "" Then
        strUpdDate = RightNow
    Else
        strUpdDate = strExchTime
    End If
    
    If strUpdTime = "" Then
        strUpdTime = strUpdDate
    End If
    If strUpdEn = "" Then
        strUpdName = GaryGi(1)
    Else
        strUpdName = strUpdEn
    End If
    intCompCode = rsData("CompCode")
    Call WriteRecordVodProcedure("2.2 進入新增控制台資料 --> " & strCMD)
    varBookMark = rsData.AbsolutePosition
    
    Select Case UCase(strCMD)
        Case "E1"
            strCmdName = "重設訂購密碼"
        Case "C4"
            strCmdName = "設定/變更區域"
        Case "B11"
            strCmdName = "訂購SVOD"
        Case "B12"
            strCmdName = "取消SVOD"
        Case "Z2"
            strCmdName = "取QSP AccountUID帳單明細"
        Case "Z3"
            strCmdName = "取QSP公司別日期範圍帳單"
        Case "Z4"
            strCmdName = "建立STB設備於QSP中"
        Case "Z5"
            strCmdName = "於QSP中移除STB設備"
    End Select
'    Dim objPbr1 As Object
'    If blnShowMsg Then
'        If frmSO1623A Is Nothing Then
'            Set objPbr1 = Nothing
'        Else
'            Set objPbr1 = frmSO1623A.pbr1
'        End If
'    Else
'        Set objPbr1 = Nothing
'    End If
    '反向即時命令如果VodAccountId是空值,則不要送命令,直接回成功 For Jacky By Kin 2014/0716
     If (IsCancelCmd) And (Len(strProcessingDate & "") = 0) And (Len(rsData("VodAccountId") & "") = 0) Then
        strResult = "1"
        SetVodCmd = True
        Call WriteRecordVodProcedure("2.2 反向即時而且VodAccountId無值，不送命令! --> " & strCMD)
        GoTo 88
    End If
    If blnExeCommand Then
        Call WriteRecordVodProcedure("2.2 進入新增控制台資料 --> " & strCMD)
        If Not objVodCommand.Vod_Command(gcnGi.ConnectionString, intCompCode & "", GetOwner, strCMD _
                                , , , , , strUpdDate, , , strServiceProviderID, strAccountNum, _
                                strAccoutName, , strAddress, strPostCode, _
                                strAccountUID, strLoginID, _
                                strPin, , , strSmartCardID, strMacAddress, _
                                strSerialNumber, , , , strCASNo _
                                , strProductID, strProductName, strUpdTime, strUpdName, _
                                strSTBUId, strProcessingDate, _
                                strSNO, strMediaBillNo, strReasonCode, strReasonName, GaryGi(0), _
                                str(intCompCode), strBillStartDate, strBillEndDate, _
                                strICCUID, strCitemStr, , , , , , strCitemCodes, , , , GetPrbCtl, rsResult, strResult, strRetErrMsg, _
                                strFaultReason) Then GoTo 88
            
        Call WriteRecordVodProcedure("2.3 完成新增控制台資料 --> " & strCMD)
    End If
    
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        SetVodCmd = True
        Call WriteRecordVodProcedure("2.4 預約資料 --> " & strCMD)
        GoTo 88
    End If
    rsData.AbsolutePosition = varBookMark
    Call WriteRecordVodProcedure("2.4.1 開始更新資料 --> " & strCMD)
    Select Case UCase(strCMD)
        Case "B11"      '以填入SO555的資料為主,填入SO555B,如果SO555B不存在則新增,存在則Upd
            Call GetRs555(rs555)
            If rs555 Is Nothing Then strRetErrMsg = "找SO555資料錯誤" & vbCrLf & "SEQNO=" & strCmdSeqNo: strResult = "0": GoTo 88
            If rs555.State <> adStateOpen Then strRetErrMsg = "SO555開啟失敗！": strResult = "0": GoTo 88
            If rs555.EOF Then strRetErrMsg = "找不到SO555 SeqNo=" & strCmdSeqNo & " 的資料！": strResult = "0": GoTo 88
            If Not UpdB11Cmd(rs555) Then strRetErrMsg = "更新SO555B失敗！": strResult = "0": GoTo 88
        Case "B12"      '以填入SO555的資料為主,更新SO555B的資料
            Call GetRs555(rs555)
            If rs555 Is Nothing Then strRetErrMsg = "找SO555資料錯誤" & vbCrLf & "SEQNO=" & strCmdSeqNo: GoTo 88
            If rs555.State <> adStateOpen Then strRetErrMsg = "rs555開啟失敗！": strResult = "0": GoTo 88
            If rs555.EOF Then strRetErrMsg = "找不到SeqNo=" & strCmdSeqNo & " 的資料！": strResult = "0": GoTo 88
            If Not UpdB12Cmd(rs555, strSNO) Then strRetErrMsg = "更新SO555B失敗！": strResult = "0": GoTo 88
        Case "Z4"
            If Not rsResult Is Nothing Then
                If rsResult.State = adStateClosed Then
                    strRetErrMsg = "找不到GateWay回填的資料！"
                    GoTo 88
                Else
                    If rsResult.EOF Then
                        strRetErrMsg = "找不到GateWay回填的資料！"
                        GoTo 88
                    Else
                        strSTBUId = rsResult("STBUID") & ""
                    End If
                End If
            Else
                strRetErrMsg = "找不到GateWay回填的資料！"
                GoTo 88
            End If
        
    End Select
    strResult = "1"
    SetVodCmd = True
    Call WriteRecordVodProcedure("2.5 完成更新資料 --> " & strCMD)
88:
    If blnShowMsg Then
        If IsA6Command Or IsRestCommand Then
'            If UCase(strCMD) <> "B11" Then
'                ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
'            End If
        Else
            ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
        End If
    End If
    On Error Resume Next
    'Call CloseRecordset(rs005)
    Call CloseRecordset(rs555)
    'Call CloseRecordset(rs005B)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "SetVod")
End Function
Private Function UpdB12Cmd(ByRef rs555 As ADODB.Recordset, Optional ByVal aPrSNo As String = Empty) As Boolean
  On Error GoTo ChkErr
    Dim rs555B As New ADODB.Recordset
    UpdB12Cmd = False
    rs555.MoveFirst
    Do While Not rs555.EOF
        Call GetRs555B(rs555B, rs555("PID"), rs555("AccountNum"), rs555("AccountUID"), False)
        If rs555B.EOF Then Exit Function
        With rs555B
            .Fields("Status") = 0
            '.Fields("PRSNo") = aPrSNo
            .Fields("CloseTime") = rs555("ExchTime")
            .Update
        End With
        rs555.MoveNext
    Loop
    UpdB12Cmd = True
    On Error Resume Next
    Call CloseRecordset(rs555B)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "UpdB12Cmd")
End Function
Private Function UpdB11Cmd(ByRef rs555 As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim aCitemCode As String
    Dim aServiceType As String
    Dim rs555B As New ADODB.Recordset
    UpdB11Cmd = False
    rs555.MoveFirst
    Do While Not rs555.EOF
        aServiceType = Empty
        aCitemCode = Empty
        Call GetRs555B(rs555B, rs555("PID") & "", rs555("AccountNum") & "", rs555("AccountUID") & "", True)
        With rs555B
            If .EOF Then
                .AddNew
                'Call GetCitemCode(rs555("PID"), aCitemCode, aServiceType)
                .Fields("VODAccountID") = rs555("AccountUID") & ""
                .Fields("ProductID") = rs555("PID") & ""
                .Fields("ProductName") = rs555("PName") & ""
                .Fields("Status") = 1
                .Fields("OpenTime") = rs555("ExchTime")
                '.Fields("CitemCode") = aCitemCode
                .Fields("CitemCode") = NoZero(rs555("CitemCode") & "")
                .Update
            Else
                .Fields("Status") = 1
                .Fields("OpenTime") = rs555("ExchTime")
                .Fields("CloseTime") = Null
                .Update
            End If
        End With
        rs555.MoveNext
    Loop
    UpdB11Cmd = True
    On Error Resume Next
    Call CloseRecordset(rs555B)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "UpdB11Cmd")
End Function

Public Sub GetCitemCode(ByVal strCode As String, ByRef aCCode As String, ByRef aService As String)
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim rsCD019A As New ADODB.Recordset
    strSQL = "Select A.CitemCode,C.ServiceType From " & GetOwner & "CD019A A," & _
            GetOwner & "CD024 B," & GetOwner & "CD019 C " & _
            " Where A.CodeNo='" & strCode & "' And A.CodeNo=B.CodeNO " & _
            " And B.VodType=2 And B.StopFlag<> 1 And A.CitemCode=C.CodeNo"
    If Not GetRS(rsCD019A, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then Exit Sub
    If Not rsCD019A.EOF Then
        aCCode = rsCD019A("CitemCode")
        aService = rsCD019A("ServiceType")
    End If
    Exit Sub
ChkErr:
  Call ErrSub(FormName, "GetCitemCode")
End Sub
Public Sub GetRs555B(ByRef rs As ADODB.Recordset, ByVal strProductID As String, _
                    ByVal aCustId As String, ByVal aVodAccountId As String, _
                    ByVal aOrder As Boolean)
  On Error GoTo ChkErr
    Dim strSQL As String
'    strSQL = "Select A.RowID,A.* From " & GetOwner & "SO555B A " & _
'            " Where CustId=" & aCustId & " And ProductID='" & strProductID & "'" & _
'            " AND VODAccountID=" & aVODAccountId & " And Status = " & IIf(aOrder, "1", "0")
            
    strSQL = "Select A.RowID,A.* From " & GetOwner & "SO555B A " & _
            " Where  ProductID='" & strProductID & "'" & _
            " AND VODAccountID=" & aVodAccountId & " And Status = " & IIf(aOrder, "0", "1")
    If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then Exit Sub
  Exit Sub
ChkErr:
  Call ErrSub(FormName, "GetRs005")
End Sub
Public Sub GetRs555(ByRef rs As ADODB.Recordset)
  On Error GoTo ChkErr
    Dim strSQL As String
    strSQL = "Select A.*,B.ProductId PID,B.ProductName PName,C.ChannelID,B.CitemCode From " & _
                strCMowner & "SO555 A," & _
                strCMowner & "SO555A B," & strCMowner & "CD024 C " & _
                " Where A.SeqNo=" & strCmdSeqNo & " And A.SeqNo=B.SeqNo " & _
                " And C.CodeNo=B.ProductId And C.VodType=2 And C.StopFlag<>1"
  
    If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then
        Set rs = Nothing
    End If
    Exit Sub
ChkErr:
  Call ErrSub(FormName, "GetRs555")
End Sub
Public Function GetSO555BCitemCode(ByVal aVodAccountId As String, Optional ByVal aStatus As Integer = 1) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim aRet As String
    Dim rsTmp As New ADODB.Recordset
    aSQL = "SELECT CitemCode FROM " & strCMowner & "SO555B " & _
            " WHERE VODAccountID = " & aVodAccountId & _
            " AND Status = " & aStatus
    If Not GetRS(rsTmp, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
    Do While Not rsTmp.EOF
        If Len(aRet) = 0 Then
            aRet = rsTmp("CitemCode") & ""
        Else
            aRet = aRet & "," & rsTmp("CitemCode") & ""
        End If
        rsTmp.MoveNext
    Loop
    GetSO555BCitemCode = aRet
    On Error Resume Next
    CloseRecordset rsTmp
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetProductID")
End Function

'5696 如果是A6命令要多傳ProductID By Kin 2010/06/17
Public Function GetProductId(ByVal aVodAccount As String, Optional ByVal aStatus As Integer = 1) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim strCodeNo As String
    Dim rsSO555B As New ADODB.Recordset
    Dim rsCD024 As New ADODB.Recordset
    Dim aRet As String
    
    aSQL = "SELECT ProductID FROM " & strCMowner & "SO555B " & _
        " WHERE NVL(Status,0)=" & aStatus & " AND VODAccountID = " & aVodAccount
    If Not GetRS(rsSO555B, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then Exit Function
    If rsSO555B.EOF Then GetProductId = Empty: Exit Function
    Do While Not rsSO555B.EOF
        If strCodeNo = "" Then
            strCodeNo = "'" & rsSO555B("ProductID") & "'"
        Else
            strCodeNo = strCodeNo & ",'" & rsSO555B("ProductID") & "'"
        End If
        rsSO555B.MoveNext
    Loop
    aSQL = "SELECT ChannelID FROM " & strCMowner & "CD024 " & _
            " WHERE CodeNo IN(" & strCodeNo & ")"
    If Not GetRS(rsCD024, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then Exit Function
    Do While Not rsCD024.EOF
        If aRet = "" Then
            aRet = rsCD024("ChannelID")
        Else
            aRet = aRet & "," & rsCD024("ChannelID")
        End If
        rsCD024.MoveNext
    Loop
    GetProductId = aRet
    On Error Resume Next
    Call CloseRecordset(rsSO555B)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "GetProductID")
End Function
Public Function QueryAccInfo(ByRef rsData As ADODB.Recordset, ByVal strCMD As String, _
                        Optional ByVal strExchTime As String = "", Optional ByVal strServiceProviderID As String = "", _
                        Optional ByVal strAccountNum As String, Optional ByVal strAccoutName As String, _
                        Optional ByVal strAddress As String = "", Optional ByVal strPostCode As String = "", _
                        Optional ByVal strLoginID As String = "", Optional ByVal strPin As String = "", _
                        Optional ByVal strSmartCardID As String = "", _
                        Optional ByVal strMacAddress As String = "", Optional ByVal strSerialNumber As String = "", _
                        Optional ByVal strCASNo As String = "", Optional ByVal strAccountUID As String, _
                        Optional ByVal strProductID As String = "", Optional ByVal strProductName As String, _
                        Optional ByVal strUpdTime As String = "", Optional ByVal strUpdEn As String = "", _
                        Optional ByRef strSTBUId As String = "", _
                        Optional ByVal strProcessingDate As String = "", Optional ByVal strSNO As String = "", _
                        Optional ByVal strMediaBillNo As String = "", Optional ByVal strReasonCode As String = "", _
                        Optional ByVal strReasonName As String = "", _
                        Optional ByVal strBillStartDate As String = "", Optional ByVal strBillEndDate As String = "", _
                        Optional ByVal strICCUID As String = "", Optional ByVal strOldSerialNumber As String = "", _
                        Optional ByVal strOldSmartCardNo As String = "", Optional strNotes As String = "", _
                        Optional strDVRQuota As String = "", Optional strTunerCount As String = "", _
                        Optional strCitemCodes As String = "", _
                        Optional ByRef rsResult As Object = Nothing, Optional strResult As String, _
                        Optional ByRef strRetErrMsg As String, Optional ByRef strFaultReason As String) As Boolean
                        
    On Error GoTo ChkErr
        Dim strCmdName As String
        Dim strUpdName As String
        Dim strUpdDate As String
        Dim intCompCode As Integer
        Dim aMsg As String
        Dim aCitemNames As String
        Dim objVodCommand As New clsVodcommand
         strCmdName = "查詢帳號及產品ID"
        If strExchTime = "" Then
            strUpdDate = RightNow
        Else
             strUpdDate = strExchTime
        End If
        If strUpdEn = "" Then
             strUpdName = GaryGi(1)
        Else
            strUpdName = strUpdEn
        End If
         intCompCode = rsData("CompCode")
        If Not objVodCommand.Vod_Command(gcnGi.ConnectionString, intCompCode & "", GetOwner, strCMD _
                                    , , , , , strUpdDate, , , strServiceProviderID, strAccountNum, _
                                    strAccoutName, , strAddress, strPostCode, strAccountUID, strLoginID, _
                                    strPin, , , strSmartCardID, strMacAddress, strSerialNumber, , , , strCASNo _
                                     , strProductID, strProductName, strUpdTime, strUpdName, _
                                     strSTBUId, strProcessingDate, _
                                     strSNO, strMediaBillNo, strReasonCode, strReasonName, GaryGi(0), _
                                     str(intCompCode), strBillStartDate, strBillEndDate, _
                                       strICCUID, , strOldSerialNumber, strOldSmartCardNo, _
                                        strNotes, strDVRQuota, strTunerCount, _
                                         strCitemCodes, , , , _
                                       GetPrbCtl, rsResult, strResult, _
                                    strRetErrMsg, strFaultReason) Then
            ShowMsg strCmdName, strResult, strFaultReason, strRetErrMsg
        Else
            If Not rsResult Is Nothing Then
                If (rsResult.RecordCount > 0) Then
                    If Len(rsResult("SMARTCARDID") & "") > 0 Then
                        aMsg = "卡號：" & rsResult("SMARTCARDID")
                    End If
                    If Len(rsResult("AccountUID") & "") > 0 Then
                        aMsg = aMsg & "VOD帳號：" & rsResult("AccountUId") & vbCrLf
                    End If
                    aCitemNames = GetProductName(rsResult("Productid"))
                    If Len(aCitemNames & "") > 0 Then
                        aMsg = aMsg & "收費項目名稱：" & aCitemNames
                    End If
                    If Len(rsResult("AccountUID") & "") > 0 Then
                        MsgBox aMsg, vbInformation, "結果"
                    Else
                        MsgBox "查無資料！", vbCritical, "警告"
                    End If
                    
                Else
                    MsgBox "查無資料！", vbCritical, "警告"
                End If
            Else
                MsgBox "查無資料！", vbCritical, "警告"
            End If
        
        End If
    
    Exit Function
ChkErr:
Call ErrSub(FormName, "OpenVod")
End Function
Private Function GetProductName(ByVal productid As String) As String
  On Error GoTo ChkErr
    If Len(productid & "") = 0 Then Exit Function
    Dim aResult As String
    Dim aSQL As String
    Dim arrayProductid() As String
    Dim aProductid As String
    Dim i As Integer
    Dim rsProduct As New ADODB.Recordset
    arrayProductid = Split(productid, ",")
    For i = LBound(arrayProductid) To UBound(arrayProductid)
        If Len(aProductid & "") = 0 Then
            aProductid = "'" & Replace(arrayProductid(i), "'", "") & "'"
        Else
            aProductid = aProductid & ",'" & Replace(arrayProductid(i), "'", "") & "'"
        End If
    Next i
    aSQL = "select distinct c.description from " & GetOwner & "cd024 a , " & GetOwner & "cd019a b ," & GetOwner & "cd019 c " & _
                " where channelid IN (" & aProductid & ") and nvl(a.stopflag,0)=0 and a.codeno=b.codeno and b.citemcode=c.codeno " & _
                " and nvl(c.stopflag,0) = 0 "
    If Not GetRS(rsProduct, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
    Do While Not rsProduct.EOF
        If Len(aResult & "") = 0 Then
            aResult = rsProduct("DESCRIPTION")
        Else
            aResult = aResult & "," & rsProduct("DESCRIPTION")
        End If
        rsProduct.MoveNext
    Loop
    On Error Resume Next
    GetProductName = aResult
    Call CloseRecordset(rsProduct)
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetProductName")
End Function

Public Function OpenVod(ByRef rsData As ADODB.Recordset, ByVal strCMD As String, _
                        Optional ByVal strExchTime As String = "", Optional ByVal strServiceProviderID As String = "", _
                        Optional ByVal strAccountNum As String, Optional ByVal strAccoutName As String, _
                        Optional ByVal strAddress As String = "", Optional ByVal strPostCode As String = "", _
                        Optional ByVal strLoginID As String = "", Optional ByVal strPin As String = "", _
                        Optional ByVal strSmartCardID As String = "", _
                        Optional ByVal strMacAddress As String = "", Optional ByVal strSerialNumber As String = "", _
                        Optional ByVal strCASNo As String = "", Optional ByVal strAccountUID As String, _
                        Optional ByVal strProductID As String = "", Optional ByVal strProductName As String, _
                        Optional ByVal strUpdTime As String = "", Optional ByVal strUpdEn As String = "", _
                        Optional ByRef strSTBUId As String = "", _
                        Optional ByVal strProcessingDate As String = "", Optional ByVal strSNO As String = "", _
                        Optional ByVal strMediaBillNo As String = "", Optional ByVal strReasonCode As String = "", _
                        Optional ByVal strReasonName As String = "", _
                        Optional ByVal strBillStartDate As String = "", Optional ByVal strBillEndDate As String = "", _
                        Optional ByVal strICCUID As String = "", Optional ByVal strOldSerialNumber As String = "", _
                        Optional ByVal strOldSmartCardNo As String = "", Optional strNotes As String = "", _
                        Optional strDVRQuota As String = "", Optional strTunerCount As String = "", _
                        Optional strCitemCodes As String = "", _
                        Optional ByRef rsResult As Object = Nothing, Optional strResult As String, _
                        Optional ByRef strRetErrMsg As String, Optional ByRef strFaultReason As String) As Boolean
  On Error GoTo ChkErr
    OpenVod = False
    Dim strUpdDate As String
    Dim strUpdName As String
    Dim varBookMark As Variant
    Dim strCmdName As String
    Dim strUPDSOAC0201B As String
    Dim rsSO004G As New ADODB.Recordset
    Dim objVodCommand As New clsVodcommand
    Dim aOLDICCUID As String
    Dim aNewICCUID As String
    Dim rsOldSO004G As New ADODB.Recordset
    
    
    If strExchTime = "" Then
        strUpdDate = RightNow
    Else
        strUpdDate = strExchTime
    End If
    If strUpdTime = "" Then
        strUpdTime = strUpdDate
    End If
    If strUpdEn = "" Then
        strUpdName = GaryGi(1)
    Else
        strUpdName = strUpdEn
    End If
    strResult = "0"
    Call WriteRecordVodProcedure("2.1 開始新增控制台資料 --> " & strCMD)
    varBookMark = rsData.AbsolutePosition
    
    Select Case UCase(strCMD)
        Case "A1"
            strCmdName = "啟動訂購權"
        Case "A2"
            strCmdName = "關閉訂購權"
        Case "A3"
            strCmdName = "暫停訂購權"
            
        Case "A4"
            strCmdName = "恢復訂購權"
        Case "A6"
            strCmdName = "維修換拆(整組換)"
'        Case "A7"
'            strCmdName = "配對"
        Case "A8"
            strCmdName = "維修換拆 STB Swap"
        Case "A9"
            strCmdName = "維修換拆 Smart Card Swap"
        Case "C1"
            strCmdName = "查詢帳號及產品ID"
    End Select
'    Dim objPbr1 As Object
'    If blnShowMsg Then
'        If frmSO1623A Is Nothing Then
'            Set objPbr1 = Nothing
'        Else
'            Set objPbr1 = frmSO1623A.pbr1
'        End If
'    Else
'        Set objPbr1 = Nothing
'    End If
'
    intCompCode = rsData("CompCode")
    '反向即時命令如果VodAccountId是空值,則不要送命令,直接回成功 For Jacky By Kin 2014/0716
    If (IsCancelCmd) And (Len(strProcessingDate & "") = 0) And (Len(rsData("VodAccountId") & "") = 0) Then
        strResult = "1"
        OpenVod = True
        Call WriteRecordVodProcedure("2.2 反向即時而且VodAccountId無值，不送命令! --> " & strCMD)
        GoTo 88
    End If
    
    If (blnExeCommand) Then
        Call WriteRecordVodProcedure("2.2 進入新增控制台資料 --> " & strCMD)
        If Not objVodCommand.Vod_Command(gcnGi.ConnectionString, intCompCode & "", GetOwner, strCMD _
                                , , , , , strUpdDate, , , strServiceProviderID, strAccountNum, _
                                strAccoutName, , strAddress, strPostCode, strAccountUID, strLoginID, _
                                strPin, , , strSmartCardID, strMacAddress, strSerialNumber, , , , strCASNo _
                                 , strProductID, strProductName, strUpdTime, strUpdName, _
                                 strSTBUId, strProcessingDate, _
                                 strSNO, strMediaBillNo, strReasonCode, strReasonName, GaryGi(0), _
                                 str(intCompCode), strBillStartDate, strBillEndDate, _
                                   strICCUID, , strOldSerialNumber, strOldSmartCardNo, _
                                    strNotes, strDVRQuota, strTunerCount, _
                                     strCitemCodes, , , , _
                                   GetPrbCtl, rsResult, strResult, _
                                strRetErrMsg, strFaultReason) Then GoTo 88
            
        Call WriteRecordVodProcedure("2.3 完成新增控制台資料 --> " & strCMD)
    End If
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        OpenVod = True
        Call WriteRecordVodProcedure("2.2 預約資料 --> " & strCMD)
        GoTo 88
    End If
    Call WriteRecordVodProcedure("2.4 進入更新資料 --> " & strCMD)
    rsData.AbsolutePosition = varBookMark
    Call WriteRecordVodProcedure("2.4.1 開始更新資料 --> " & strCMD)
    Select Case UCase(strCMD)
        Case "A1"
            If rsResult Is Nothing Then
                strRetErrMsg = "無回填VOD帳號"
                strResult = "0"
                Call WriteRecordVodProcedure("-666 無回填VOD帳號 ")
                GoTo 88
            End If
            
            If Not GetRS(rsSO004G, "Select * From " & GetOwner & "SO004G Where " & _
                        " 1=0 ", gcnGi, _
                        adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then GoTo 88
            
            rsSO004G.AddNew
            rsSO004G("VODCredit") = Val(GetRsValue("Select VODCredit From " & GetOwner & "SO041 Where SysId=" & intCompCode, gcnGi) & "")
            rsSO004G("UpdEn") = strUpdName
            rsSO004G("UpdTime") = GetDTString(strUpdDate)
            rsSO004G("PrePay") = 0
            rsSO004G("Present") = 0
            rsSO004G("SendCount") = 0
            rsSO004G("VODAccountId") = rsResult("AccountUID").Value
            rsSO004G("MVODID") = rsResult("AccountUID").Value
            rsSO004G("VodType") = GetVodType(rsData("ModelCode") & "")
            '#6617 要將舊設備的SO004G複制一份至新的設備上 By Kin 2013/10/01
            '#7068 要增加VODType By Kin 2015/08/06
            If Len(OldVodAccountId) > 0 Then
                If Not GetRS(rsOldSO004G, "SELECT * FROM " & GetOwner & "SO004G " & _
                    " WHERE VODACCOUNTID = " & OldVodAccountId, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then GoTo 88
                If rsOldSO004G.RecordCount > 0 Then
                    rsSO004G("PrePay") = rsOldSO004G("PrePay")
                    rsSO004G("Present") = rsOldSO004G("Present")
                    rsSO004G("VODCredit") = rsOldSO004G("VODCredit")
                    rsSO004G("SalePointcode") = rsOldSO004G("SalePointcode")
                    rsSO004G("SalePointname") = rsOldSO004G("SalePointname")
                    rsSO004G("SendCount") = rsOldSO004G("SendCount")
                    rsSO004G("NoPayCredit") = rsOldSO004G("NoPayCredit")
                    rsSO004G("UpdEn") = rsOldSO004G("UpdEn")
                    rsSO004G("UpdTime") = rsOldSO004G("UpdTime")
                    rsSO004G("PrintDetail") = rsOldSO004G("PrintDetail")
                    rsSO004G("StopType") = rsOldSO004G("StopType")
                End If
            End If
            rsSO004G.Update
            rsData("VODStatus") = 1
            rsData("VodAccountId") = rsResult("AccountUID").Value
            '#5603 要將SO555.ICCUID回填至SOAC0201B By Kin 2010/03/29
            If Val(rsResult("ICCUID") & "") > 0 Then
                strUPDSOAC0201B = "UpDate " & GetOwner & "SOAC0201B " & _
                    " SET ICCUID=" & rsResult("ICCUID") & _
                    " WHERE FACISNO='" & strSmartCardID & "'"
                gcnGi.Execute strUPDSOAC0201B
            End If
            
            If Len(OldVodAccountId) > 0 Then
                gcnGi.Execute "Update " & GetOwner & "SO182 " & _
                                    " SET VODAccountID = " & rsResult("AccountUID") & _
                                    " WHERE VODACCOUNTID = " & OldVodAccountId & _
                                    " AND COMPCODE = " & intCompCode
                
                 gcnGi.Execute "Update " & GetOwner & "SO033VOD " & _
                                    " SET VODAccountID = " & rsResult("AccountUID") & _
                                    " WHERE VODACCOUNTID = " & OldVodAccountId & _
                                    " AND COMPCODE = " & intCompCode
                
                gcnGi.Execute "Update " & GetOwner & "SO555B " & _
                                    " SET VODAccountID = " & rsResult("AccountUID") & _
                                    " WHERE VODACCOUNTID = " & OldVodAccountId
                
            End If
        Case "A2"
            '#5722 送命令的時候SO555要填入SOAC0201B.ICCUID,成功後要清空SOAC0201B.ICCUID By Kin 2010/07/21
            rsData("VODStatus") = 0
            strUPDSOAC0201B = "UPDATE " & GetOwner & "SOAC0201B " & _
                " SET ICCUID = NULL " & _
                " WHERE FaciSNo = '" & strSmartCardID & "' " & _
                " AND COMPCODE = " & intCompCode
            gcnGi.Execute strUPDSOAC0201B
        Case "A3"
            rsData("VODStatus") = 2
        Case "A4"
            rsData("VODStatus") = 1
            If Not GetRS(rsSO004G, "Select StopType From " & GetOwner & "SO004G Where " & _
                " VODAccountID = " & rsData("VODAccountID"), gcnGi, _
                adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then GoTo 88
            If Not rsSO004G.EOF Then
                rsSO004G.Fields(0) = 0
                rsSO004G.Update
            End If
        '#5469 如果是維換要將舊設備改為0 By Kin 2010/01/07
        Case "A6"
            '#5497 要將新設備的狀態改成舊設備的狀態 By Kin 2010/03/04
            '#5277 要將舊的smartcardno的SOAC0201B.ICCUID清空,將新的smartcardno的Soac0201B.ICCUID值 upd成新的 By Kin 2010/07/21
            rs2("VodStatus") = Val(rsData("VODStatus") & "")
            rs2.Update
            rsData("VODStatus") = 0
            If Not rsResult Is Nothing Then
                If rsResult("ICCUID") & "" <> "" Then
                    If InStr(1, rsResult("ICCUID"), ",") > 0 Then
                        aOLDICCUID = Mid(rsResult("ICCUID"), 1, InStr(1, rsResult("ICCUID"), ",") - 1)
                        aNewICCUID = Mid(rsResult("ICCUID"), InStr(1, rsResult("ICCUID"), ",") + 1)
                        'aOLDICCUID = Split(rsResult("ICCUID"), ",")(1)
                        If aOLDICCUID <> "" Then
                            strUPDSOAC0201B = "UPDATE " & GetOwner & "SOAC0201B " & _
                                " SET ICCUID = NULL " & _
                                " WHERE FACISNO ='" & Split(strSmartCardID, ",")(0) & "'" & _
                                " AND COMPCODE = " & intCompCode
                            gcnGi.Execute strUPDSOAC0201B
                        End If
                        If aNewICCUID <> "" Then
                            strUPDSOAC0201B = "UPDATE " & GetOwner & "SOAC0201B " & _
                                " SET ICCUID = " & aNewICCUID & _
                                " WHERE FACISNO = '" & Split(strSmartCardID, ",")(1) & "'" & _
                                " AND COMPCODE = " & intCompCode
                            gcnGi.Execute strUPDSOAC0201B
                        End If
                    End If
                    
                End If
            End If
        Case "C1"
            
    End Select
    rsData("UpdEn") = strUpdName
    rsData("UpdTime") = GetDTString(strUpdDate)
    rsData.Update
    Call WriteRecordVodProcedure("2.5 更新資料完成 --> " & strCMD)
    strResult = "1"
    OpenVod = True
88:
    '#6140 A1、A2特定錯誤要做返向指令 By Kin 2011/11/08
    Select Case UCase(strCMD)
        Case "A1"
            blnReturnCmd = InStr(1, strRetErrMsg, strA1ErrCode) > 0
            If (intRunRetunCount > 0) And (strResult = "1") Then    '執行第一次返向指令
                blnReturnCmd = True
            End If
            If (Not blnReturnCmd) And (intRunRetunCount = 0) Then
                intRunRetunCount = 99
            End If
        Case "A2"
            blnReturnCmd = InStr(1, strRetErrMsg, strA2ErrCode) > 0
            If (intRunRetunCount > 0) And (strResult = "1") Then    '執行第一次返向指令
                blnReturnCmd = True
            End If
            If (Not blnReturnCmd) And (intRunRetunCount = 0) Then
                intRunRetunCount = 99
            End If
        Case Else
            blnReturnCmd = False
            intRunRetunCount = 99
    End Select
        
    If (intRunRetunCount > 0) And (strResult <> "1") Then '表示已經跑過一次,如果錯就直接回錯誤了
        blnReturnCmd = False
        intRunRetunCount = 99
    End If
    If (strResult = "1") And (Not blnReturnCmd) Then intRunRetunCount = 99
    If (intRunRetunCount >= 2) Then blnReturnCmd = False: intRunRetunCount = 99 '反向指令都跑完了,不要再跑下去了 By Kin 2011/11/08
    
    If IsA6Command Then
        blnReturnCmd = False
        intRunRetunCount = 99
    End If
    
    If (Not blnReturnCmd) And (intRunRetunCount > 0) Then
        intRunRetunCount = 0
        '#6573 A6命令拆成2筆,所以判斷那個命令要ShowMsg By Kin 2013/08/27
        If IsA6Command Or IsRestCommand Then
'            If (strCMD = UCase("A1")) Then
'                ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
'            Else
'                If strResult <> "1" Then
'                    ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
'                End If
'            End If
        Else
            If blnShowMsg Then
                ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
            End If
        End If
        
    Else
        intRunRetunCount = intRunRetunCount + 1
    End If
    On Error Resume Next
    
    Call CloseRecordset(rsSO004G)
    Call CloseRecordset(rsOldSO004G)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "OpenVod")
End Function
Public Function GetVodType(ByVal ModelCode As String) As Integer
  On Error GoTo ChkErr
    Dim rsVodType As New ADODB.Recordset
    Dim aSQL As String
    GetVodType = 0
    aSQL = "Select Nvl(VodType,0) VodType From " & GetOwner & "CD043 Where CodeNo = " & ModelCode
    If Len(ModelCode) > 0 Then
        If Not GetRS(rsVodType, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        If rsVodType.RecordCount > 0 Then
            GetVodType = Val(rsVodType("VodType") & "")
        End If
    End If
    On Error Resume Next
    Call CloseRecordset(rsVodType)
  Exit Function
ChkErr:
 Call ErrSub(FormName, "GetVodType")
End Function
Public Function GetICCUID(ByVal aSmartCardNo As String, ByVal aCompCode As String) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim aRet As String
    aSQL = "SELECT ICCUID FROM " & GetOwner & "SOAC0201B " & _
        " WHERE FaciSNo = '" & aSmartCardNo & "' " & _
        " AND COMPCODE = " & aCompCode
    
    If aSmartCardNo = "" Then
        GetICCUID = ""
        Exit Function
    End If
    aRet = GetRsValue(aSQL, gcnGi) & ""
    GetICCUID = aRet
    Exit Function
ChkErr:
  Call ErrSub(FormName, "GetICCUID")
End Function
Private Function TurnNum(ByVal strChgNum) As Long
  On Error GoTo ChkErr
    Dim i As Long
    Dim strNum As String
    For i = 0 To Len(strChgNum) - 1
        If IsNumeric(Mid(strChgNum, i + 1, 1)) Then
            strNum = strNum & Mid(strChgNum, i + 1, 1)
        End If
    Next i
    TurnNum = Val(strNum)
  Exit Function
ChkErr:
    ErrSub FormName, "TurnNum"
End Function




Public Function GetStateData(objCn, ByVal intCompCode As Integer) As String
  On Error Resume Next
    GetStateData = objCn.Execute("SELECT CmdStatus FROM " & strCMowner & "SO555" & _
                                                        " WHERE  SEQNO='" & strCmdSeqNo & "'")(0) & ""
                                                        
                                                
End Function
Public Function GetErrMsg(objCn, ByVal intCompCode As Integer, ByRef strErrMsg As String, ByRef strFaultReason As String) As Boolean
    On Error Resume Next
    'Dim varErrPara As String
    GetErrMsg = False
    Dim rsErr As New ADODB.Recordset
    
'    If Not GetRS(rsErr, "Select RetMsg,SysMsg From " & strCMowner & "SO555" & _
'                                                " Where  SEQNO='" & strCmdSeqNo & "'", objCn, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'    varErrPara = objCn.Execute("Select RetMsg,SysMsg From " & strCMowner & "SO555" & _
'                                                " Where  SEQNO='" & strCmdSeqNo & "'").GetString(, , ",")
    '#6140 測試不OK,因為ErrMsg有包含逗點,所以用Split(,)會有問題
    Set rsErr = objCn.Execute("Select RetMsg,SysMsg From " & strCMowner & "SO555" & _
                                                " Where  SEQNO='" & strCmdSeqNo & "'")
                                                
    strErrMsg = Replace(rsErr("RetMsg"), Chr(13), "")
    strFaultReason = Replace(rsErr("SysMsg"), Chr(13), "")
    'strErrMsg = Split(Replace(varErrPara, Chr(13), ""), ",")(1)
    'strFaultReason = Split(varErrPara, ",")(0)
    If (strErrMsg = "") And (strFaultReason <> "") Then
        strErrMsg = strFaultReason
        strFaultReason = ""
    End If
    'If strErrMsg <> "" Then strErrMsg = "錯誤代碼:" & strErrMsg
    If Trim(strErrMsg) <> "" Or Trim(strFaultReason) <> "" Then
        GetErrMsg = True
    End If
    CloseRecordset rsErr
End Function
Public Sub ShowMsg(strCMD As String, strResult As String, strFaultReason As String, strErrMsg As String)
    On Error GoTo ChkErr
        If strResult = "1" Then
            MsgBox "[ " & strCMD & " ] 控制訊號已送出完成 !! ", vbInformation, gimsgPrompt
        Else
            If strFaultReason = Empty Then
                MsgBox "[ " & strCMD & " ] 失敗 !! " & vbCrLf & vbCrLf & _
                                "錯誤原因 : " & strErrMsg, vbInformation, gimsgPrompt
            Else
                MsgBox "[ " & strCMD & " ] 失敗 !! " & vbCrLf & vbCrLf & _
                                "錯誤原因 : " & strFaultReason & vbCrLf & vbCrLf & _
                                strErrMsg, vbInformation, gimsgPrompt
            End If
        End If
    Exit Sub
ChkErr:
    ErrSub FormName, "subShowMessage"
End Sub
Public Function GetClassID(ByRef rs As ADODB.Recordset, Optional strBPcode As String = "") As String
    On Error Resume Next
        Dim strClassIdFld As String
        strClassIdFld = IIf(IsCMCP(rs), "EMCCMCP", "EMCCM")
        If strBPcode = Empty Then
            If rs("BPcode") & "" = "" Then Exit Function
            GetClassID = GetRsValue("SELECT " & strClassIdFld & " FROM " & GetOwner & "CD078D" & _
                                                        " WHERE BPCODE='" & rs("BPcode") & "'" & _
                                                        " AND SERVICETYPE='" & rs("ServiceType") & "'")
        Else
            GetClassID = GetRsValue("SELECT " & strClassIdFld & " FROM " & GetOwner & "CD078D" & _
                                                        " WHERE BPCODE = '" & strBPcode & "'" & _
                                                        " AND SERVICETYPE='" & rs("ServiceType") & "'")
        End If
        If Err.Number <> 0 Then GetClassID = Empty
End Function
Public Function IsCMCP(ByRef rs As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
        IsCMCP = (GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                            " WHERE REFACISNO='" & rs("FaciSno") & "'" & _
                                            " AND COMPCODE=" & rs("CompCode") & _
                                            " AND CUSTID=" & rs("CustID") & _
                                            " AND SERVICETYPE='P'" & _
                                            " AND PRDATE IS NULL") > 0)

  Exit Function
ChkErr:
    ErrSub FormName, "IsCMCP"
End Function
Public Function IsEMTA(ByVal strFaciCode As String) As Boolean
  On Error Resume Next
    Dim rs As New ADODB.Recordset
    IsEMTA = False
    Set rs = gcnGi.Execute("Select REFNO From " & GetOwner & "CD022 Where CODENO=" & strFaciCode)
    If Not rs.EOF Then
        If rs(0) = 5 Then IsEMTA = True
    End If
    CloseRecordset rs
    
End Function
Public Function GetInstAddNo(ByRef rs As ADODB.Recordset) As String
  On Error GoTo ChkErr
    Dim strAddNoSQL As String
    Dim strReInsAddrNoSQL As String
    Dim strRet As String
    If rs.EOF Then Exit Function
    '#5149 增加檢查 是否有CATV移機未完工的移裝地址可用 By Kin 2009/06/18
    strReInsAddrNoSQL = "Select ReInstAddrNo From " & GetOwner & "SO009 " & _
                " Where PRCODE IN (SELECT CODENO FROM CD007 WHERE REFNO = 3) " & _
                " And ServiceType='C'" & _
                " AND FINTIME IS NULL AND RETURNCODE IS NULL AND RETURNDESCCODE IS NULL " & _
                " AND CUSTID =" & rs("CustId")
    strRet = GetRsValue(strReInsAddrNoSQL, gcnGi) & ""
    If strRet = "" Then
        strAddNoSQL = "Select InstAddrNo From " & GetOwner & "SO001 Where Custid=" & rs("Custid") & " And CompCode=" & rs("CompCode")
        strRet = GetRsValue(strAddNoSQL, gcnGi) & ""
    End If
    GetInstAddNo = strRet & ""
    Exit Function
ChkErr:
  ErrSub FormName, "GetInstAddNo"
End Function
'**********************************************************************************************************
'Corey 寫的Function
'利用傳入的ADDRNO(一定要裝機地址)找出SO014.ADDSORT，並且SO017C的涵蓋地址有包含此ADDRSORT的SNO列出來，
'並且利用SNO找出SO017B.SwitchID回傳，用逗號隔開每個SwitchID
Public Function GetSwitchID(ByVal strInstAddrNO As String, ByRef strErrString As String) As String
On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strAddrSort As String
    Dim strSNO As String
    Dim strSwitchID As String
    GetSwitchID = ""
    strSwitchID = ""
    strErrString = ""
    If Not GetRS(rsTmp, "Select AddrSort,FTTxFlag From " & GetOwner & "SO014 Where AddrNO=" & strInstAddrNO, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic, , , , False) Then Exit Function
    If rsTmp.RecordCount = 0 Then
        strErrString = "找不到對應地址"
        Exit Function '找不到對應的地址 Addrsort
    End If
'    If Val(rsTmp("FTTxFlag") & "") = 0 Then
'        strErrString = "該地址不在 FTTX 涵蓋範圍"
'        Exit Function
'    End If

    strAddrSort = Mid(rsTmp("AddrSort") & "", 1, GetFldMaxLen("SO017C", "AddrSortA"))
    If Not GetRS(rsTmp, "Select * From " & GetOwner & "SO017C Where AddrSortA <= '" & strAddrSort & "' and AddrSortB >= '" & strAddrSort & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then Exit Function
    '合乎條件的數量大於0則開始串語法，否則回傳空字串
    If rsTmp.RecordCount > 0 Then
        '先取得SO017C的SNO，再利用SNO找出SO017B的SwitchID
        rsTmp.MoveFirst
        With rsTmp
            Do While Not .EOF
                strSNO = GetRsValue("Select SwitchID From " & GetOwner & "SO017B where SNO=" & rsTmp("SNO") & "") & ""
                strSwitchID = strSwitchID & "," & Trim(strSNO)
                .MoveNext
            Loop
            strSwitchID = Mid(strSwitchID, 2)
        End With
        GetSwitchID = Trim(strSwitchID)
        If strSwitchID = "" Then strErrString = "找不到對應的SwitchID"
    Else
         strErrString = "找不到對應的SwitchID"
    End If
    On Error Resume Next
    CloseRecordset rsTmp
Exit Function
ChkErr:
    ErrSub FormName, "GetSwitchID"
End Function
Public Function GetFldMaxLen(ByVal uTableName As String, ByVal uColumn As String) As Double

  On Error GoTo ChkErr
  
    '此Function是取出資料庫內本身的SCHEMA的欄位大小值
    '並且SCHEMA的DATA_TYPE=NUMBER，並計算出NUMBER最大值回傳Double
    '負數的時候是從 -1.79769313486232E308 到 -4.94065645841247E-324
    '正數的時候是從 4.94065645841247E-324 到 1.79769313486232E308
    
    'uTableName=傳入Table名稱  uColumn=傳入欄位名稱

    Dim rsColumn As New ADODB.Recordset
    Dim strSQL As String
    
    GetFldMaxLen = 0

    '沒有傳資料進來則直接離開，否則都語法時會出錯誤訊息
    If IsNull(uTableName) Or IsNull(uColumn) Then Exit Function
    
    strSQL = "Select * From " & "User_Tab_Columns" & _
                " Where Table_Name='" & UCase(uTableName) & "'" & _
                " And Column_Name='" & UCase(uColumn) & "'"
    
    If Not GetRS(rsColumn, strSQL, , , , , , , , False) Then Exit Function
    
    Select Case UCase(rsColumn("Data_Type") & "")
        Case "NUMBER"
            GetFldMaxLen = (10 ^ rsColumn("Data_Precision")) - 1
            If rsColumn("Data_Scale") > 0 Then '判斷小數點是否有值
                GetFldMaxLen = GetFldMaxLen + (0.1 ^ rsColumn("Data_Scale")) * ((10 ^ rsColumn("Data_Scale")) - 1)
            End If
        Case "VARCHAR2", "VARCHAR", "CHAR", "CLOB", "BLOB", "ROWID"
            GetFldMaxLen = Val(rsColumn("Data_Length") & "")
        Case "LONG", "DATE", "UNDEFINED"
            GetFldMaxLen = 0
    End Select
    
    Call CloseRecordset(rsColumn)

  Exit Function
ChkErr:
    ErrSub FormName, "GetFldMaxLen"
End Function
'*******************************************************************************************
Public Function chkUseFlag(ByVal strFaciSNo As String, ByVal intComp As Integer) As Boolean
  On Error GoTo ChkErr:
    Dim rs As New ADODB.Recordset
    chkUseFlag = False
    'If blnUseFttx Then chkUseFlag = True: Exit Function
    
    If Not GetRS(rs, "Select Useflag From " & GetOwner & "SO306 Where HFCMAC='" & strFaciSNo & "'") Then Exit Function
    If rs.EOF Then Call CloseRecordset(rs): Exit Function
    '******************************************************************************************************
    '#3755 如果Linkmss為0不檢查UseFlag By Kin 2008/2/01
    If Val(GetRsValue("Select Nvl(Linkmss,0) From " & GetOwner & "CD039 Where CodeNo=" & intComp)) = 0 Then
        On Error Resume Next
        chkUseFlag = True
        Call CloseRecordset(rs)
        Exit Function
    End If
    '******************************************************************************************************
    If rs(0).Value <> 1 Then Call CloseRecordset(rs): Exit Function
    chkUseFlag = True
    Call CloseRecordset(rs)
    Exit Function
ChkErr:
    ErrSub FormName, "chkUseFlag"
End Function

Public Sub GetScmData(ByRef objCn As Object, ByRef rsResult As Object, ByVal strFldName As String)
  On Error Resume Next
    Dim strScmData As String
    'Dim strFldName As String
    
    strScmData = objCn.Execute("SELECT " & strFldName & " FROM " & strCMowner & "SO555" & _
                                           " WHERE  SEQNO='" & strCmdSeqNo & "'").GetString(, , ";") & ""
    
    If Right(strScmData, 1) = ";" Then
        strScmData = Mid(strScmData, 1, Len(strScmData) - 1)
    End If
    
    If strScmData <> Empty Then
     '   strFldName = "AccountUID"
        Call SetRsMemory(rsResult, strFldName, strScmData)
    Else
        Set rsResult = Nothing
    End If
End Sub
Public Sub SetRsMemory(ByRef rs As Object, ByVal strFldName As String, ByVal strFldValue As String)
  On Error GoTo ChkErr
    Dim i As Long
    Dim strAryFldName() As String
    Dim strAryFldValue() As String
    Set rs = CreateObject("ADODB.Recordset")
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    strAryFldName = Split(strFldName, ",")
    If Left(strFldValue, 1) = "," Then
        strFldValue = Chr(0) & Right(strFldValue, Len(strFldValue) - 1)
    End If
    strAryFldValue = Split(strFldValue, ";")
    For i = LBound(strAryFldName) To UBound(strAryFldName)
        rs.Fields.Append strAryFldName(i), adBSTR, 100
    Next i
    rs.Open
    rs.AddNew
    For i = LBound(strAryFldName) To UBound(strAryFldName)
        If Left(strAryFldValue(i), 1) = Chr(0) Then
            strAryFldValue(i) = "," & Right(strAryFldValue(i), Len(strAryFldValue(i)) - 1)
        End If
        rs.Fields(i).Value = strAryFldValue(i) & ""
    Next i
    rs.Update
    Exit Sub
ChkErr:
    ErrSub FormName, "SetRsMemory"
End Sub

Public Function GetVendCode(ByVal intComp As Integer) As String
  On Error Resume Next
    GetVendCode = gcnGi.Execute("Select VendorCode From " & GetOwner & "CM006 Where " & _
                                "CompCode=" & intComp & " And StopFlag=0")(0)
                        

End Function
Private Function UpdSO004F(ByRef rs As Recordset, ByVal strUpdTime As String) As Boolean
  On Error GoTo ChkErr
    Dim strQrySQL As String
    Dim i As Integer
    Dim strDialAccount, strExtendAccount, strExtendPassword As String
    Dim rsQry As New ADODB.Recordset
    If rs(2).Value = "" Then
        UpdSO004F = True
        Exit Function
    End If
    strQrySQL = "Select * From " & GetOwner & "SO004F" & _
                " Where ExtendAccount='" & rs("FixIPAccountId") & "'" & _
                " And ExtendType=1 And PrDate is Null"
    If Not GetRS(rsQry, strQrySQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsQry.EOF Then
        strDialAccount = rsQry("DialAccount") & ""
        strExtendAccount = rsQry("ExtendAccount") & ""
        strExtendPassword = rsQry("ExtendPassword") & ""
        rsQry.MoveFirst
        Do While Not rsQry.EOF
            rsQry("PrDate") = Format(strUpdTime, "YYYY/MM/DD HH:MM:SS")
            rsQry("UpdEn") = GaryGi(1)
            rsQry("UpdTime") = GetDTString(strUpdTime)
            rsQry.Update
            rsQry.MoveNext
        Loop
    End If
    For i = 0 To rs.Fields.Count - 3
        With rsQry
            .AddNew
            .Fields("SeqNo") = Get_SO004F_Seq_No(gcnGi)
            .Fields("DialAccount") = strDialAccount
            .Fields("ExtendAccount") = strExtendAccount
            .Fields("ExtendPassword") = strExtendPassword
            .Fields("FixIPAddress") = rs(i + 2).Value & ""
            .Fields("UpdEn") = GaryGi(1)
            .Fields("UpdTime") = GetDTString(strUpdTime)
            .Update
        End With
    Next i
    UpdSO004F = True
    On Error Resume Next
    Call CloseRecordset(rsQry)
    
    Exit Function
ChkErr:
  Call ErrSub(FormName, "UpdSO004F")
End Function
