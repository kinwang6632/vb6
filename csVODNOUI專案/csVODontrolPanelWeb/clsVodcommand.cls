VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVodcommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private lngCMtimeout As Long
'Private intCompCode As Integer

Public Function Vod_Command(ByVal strCn As String, ByVal strComp_Code As String, ByVal strOwner_Name As String, _
                            ByVal strCMD As String, Optional strAction As String = "", _
                            Optional ByVal strCmdStatus As String = "W", Optional ByVal strRequestTime As String = "", _
                            Optional ByVal strResponseTime As String = "", _
                            Optional ByVal strExchTime As String = "", Optional ByVal strRetMsg As String = "", _
                            Optional ByVal strSysMsg As String = "", Optional ByVal strServiceProviderID As String = "", _
                            Optional ByVal strAccountNum As String, Optional ByVal strAccoutName As String = "", _
                            Optional ByVal strAccountStatus As String = "", Optional ByVal strAddress As String = "", _
                            Optional ByVal strPostCode As String, Optional ByVal strAccountUID As String = Empty, _
                            Optional ByVal strLoginID As String = "", Optional ByVal strPin As String = "", _
                            Optional ByVal intDefaultUser As Integer = 1, Optional ByVal intAccessPointUID As Integer = 1, _
                            Optional ByVal strSmartCardID As String = "", Optional ByVal strMacAddress As String, _
                            Optional ByVal strSerialNumber As String = "", Optional ByVal strSTBstatus As String = "", _
                            Optional ByVal strDeviceType As String = "STB", Optional ByVal intDeviceEnabled As Integer, _
                            Optional ByVal strCASNo As String = "", Optional ByVal strProductID As String = Empty, _
                            Optional ByVal strProductName As String = "", Optional ByVal strUpdTime As String, Optional ByVal strUpdEn As String, _
                            Optional ByVal strSTBUId As String = "", Optional ByVal strProcessingDate As String = "", Optional ByVal strSNO As String = "", _
                            Optional ByVal strMediaBillNo As String = "", Optional ByVal strReasonCode As String = "", _
                            Optional ByVal strReasonName As String = "", Optional ByVal strOperatorId As String = "", _
                            Optional ByVal strCompCode As String = "", _
                            Optional ByVal strBillStartDate As String = "", Optional ByVal strBillEndDate As String = "", _
                            Optional ByRef objPBG As Object = Nothing, Optional ByRef rsResult As Object = Nothing, _
                            Optional ByRef strResult As String = "", Optional ByRef strRetErrMsg As String = "", _
                            Optional ByRef strFaultReason As String = "", Optional ByVal strICCUID As String = "", _
                            Optional ByVal strCitemStr As String = "") As Boolean
  On Error GoTo chkErr
    Dim str_CmdSeqNo As String
    Dim strCmdName As String
    Dim objCn As Object
    strCmdName = Empty
    Call WriteRecordVodProcedure("2.2.1 新增控制台資料 --> " & strCMD)
    Select Case UCase(strCMD)
        Case "A1"
            strCmdName = "開機"
        Case "A2"
            strCmdName = "拆機"
        Case "A3"
            strCmdName = "停機"
        Case "A4"
            strCmdName = "復機"
        Case "A6"
            strCmdName = "維修換拆(整組換)"
        Case "A7"
            strCmdName = "配對"
        Case "A8"
            strCmdName = "維修換拆"
        Case "A9"
            strCmdName = "維修換拆"
        Case "E1"
            strCmdName = "設定使用者及密碼"
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
    Vod_Command = False
    If strCmdName = "" Then strRetErrMsg = "找不到命令 !": GoTo 88
    Set objCn = CreateObject("ADODB.Connection")
    With objCn
        .CursorLocation = 3
        .Open strCn
    End With
    If objCn.State <= 0 Then
        strRetErrMsg = "資料庫連線尚未開啟 !"
        Exit Function
    End If
    strOwner = strOwner_Name
    intCompCode = strComp_Code
    On Error Resume Next
    Dim varSysPara As Variant
    Dim strQryPara As String
    strQryPara = "SELECT INTEGRATESPM,CMTIMEOUT,CMOWNER FROM " & strOwner & "SO041" & _
                                        " WHERE COMPCODE=" & intCompCode

    varSysPara = objCn.Execute(strQryPara).GetRows(1, 1)

    If Val(varSysPara(0, 0) & "") = 0 Then
        strRetErrMsg = "[SO041] 系統業者參數檔 [啟動CMCP介接] 尚未開啟, 因此並未呼叫 Web Service 介接 !"
        Exit Function
    End If
    lngCMtimeout = Val(varSysPara(1, 0) & "")
    If lngCMtimeout = 0 Then lngCMtimeout = 30
    If lngCDRTimeOut > 0 Then lngCMtimeout = lngCDRTimeOut
    strCMowner = varSysPara(2, 0) & ""
    'strCMowner = GetOwner
    If strCMowner = Empty Then
        strCMowner = ""
    Else
        strCMowner = strCMowner & "."
    End If
    If Err.Number <> 0 Then
        Err.Clear
        strRetErrMsg = "[SO041] 系統業者參數檔 查無 [公司別] 為 " & intCompCode & " 的資料 , 請確認 !"
        Exit Function
    End If
    strCmdSeqNo = Empty
    If strCmdSeqNo = Empty Then str_CmdSeqNo = Get_Cmd_Seq_No(objCn)
    strCmdSeqNo = str_CmdSeqNo
    If Not Insert_SO555_Para(objCn, strRetErrMsg, strCMD, strAction, strCmdStatus, strRequestTime, strResponseTime, _
                    strExchTime, strRetMsg, strSysMsg, strServiceProviderID, strAccountNum, strAccoutName, _
                    strAccountStatus, strAddress, strPostCode, strAccountUID, strLoginID, strPin, _
                    intDefaultUser, intAccessPointUID, _
                    strSmartCardID, strMacAddress, strSerialNumber, strSTBstatus, strDeviceType, _
                    intDeviceEnabled, strCASNo, strProductID, strProductName, _
                    strUpdTime, strUpdEn, strProcessingDate, strSNO, strMediaBillNo, _
                    strReasonCode, strReasonName, strOperatorId, strSTBUId, _
                    strCompCode, strBillStartDate, strBillEndDate, strICCUID, strCitemStr) Then GoTo 88
    If strProcessingDate <> Empty Then
        Vod_Command = True
        GoTo 88
    End If
    
    If Not Pooling(objCn, strAccountNum, strRetErrMsg, strFaultReason, objPBG) Then GoTo 88
    Call WriteRecordVodProcedure("999.Pooling結束")
    Dim aFldName As String
    Select Case UCase(strCMD)
        Case "A1"
            aFldName = "AccountUID,ICCUID"
        Case "Z4"
            aFldName = "STBUID"
        Case "A6"
            aFldName = "ICCUID"
    End Select
    If Not rsResult Is Nothing Then Call GetScmData(objCn, rsResult, aFldName)
    Vod_Command = True

88:
    On Error Resume Next
    objCn.Close
    Set objCn = Nothing
    Exit Function
chkErr:

End Function

'Public Function SCM_Command(ByVal strCn As String, ByVal strComp_Code As String, ByVal strOwner_Name As String, _
'                            ByVal strMsgCode As String, Optional ByVal strOperatorId As String = Empty, Optional strCustID As String = Empty, _
'                            Optional ByVal strCustName As String = Empty, Optional strCustTel As String = Empty, _
'                            Optional ByVal strCustAddress As String = Empty, Optional ByVal strDialAccount As String = Empty, _
'                            Optional ByVal strSchemaCode1 As String = Empty, Optional ByVal strSchemaCode2 As String = Empty, _
'                            Optional ByVal strUserIdNo As String = Empty, Optional ByVal strUserName As String = Empty, _
'                            Optional ByVal strUserBirthday As String = Empty, Optional ByVal strDeviceSNo1 As String = Empty, _
'                            Optional ByVal strDeviceSNo2 As String = Empty, Optional ByVal strDeviceType As String = Empty, _
'                            Optional ByVal strDeviceModelNo As String = Empty, Optional ByVal strPassword As String = Empty, _
'                            Optional ByVal strSNO As String = Empty, Optional ByVal strMediaBillNo As String = Empty, _
'                            Optional ByVal strProcessingDate As String = Empty, _
'                            Optional ByVal strReasonCode As String = Empty, Optional ByVal strReasonName As String = Empty, _
'                            Optional ByRef objPBG As Object = Nothing, Optional ByRef rsResult As Object = Nothing, _
'                            Optional ByRef strResult As String = Empty, Optional ByRef strFaultReason As String = Empty, _
'                            Optional ByRef strErrMsg As String = Empty, Optional ByVal strParamData As String = Empty, _
'                            Optional ByVal strDynIpCount As String = Empty, Optional ByVal strFixIPCount As String = Empty) As Boolean
'  On Error GoTo ChkErr
'    Dim str_CmdSeqNo As String
'    Dim strCmdName As String
'    Dim objCn As Object
'    strCmdName = Empty
'    Select Case strMsgCode
'        Case "10"
'            If blnUseFttx Then
'                strCmdName = "啟用PPPOE帳號"
'            Else
'                strCmdName = "建立CM服務帳號"
'            End If
'        Case "11"
'            strCmdName = "開機"
'        Case "12"
'            strCmdName = "關機"
'        Case "13"
'            strCmdName = "停機"
'        Case "14"
'            strCmdName = "復機"
'        '#4276 增加此共用指令 By Kin 2009/01/08
'        Case "15"
'            strCmdName = "終止服務帳號"
'        Case "16"
'            strCmdName = "設備更換"
'        Case "17"
'            strCmdName = "重置設備"
'        Case "18"
'            strCmdName = "預約FTTX裝機"
'        Case "19"
'            strCmdName = "FTTX裝機退單"
'        Case "20"
'            strCmdName = "變更基本資料"
'        Case "21"
'            strCmdName = "變更申請人"
'        Case "22"
'            strCmdName = "變更速率"
'        Case "23"
'            strCmdName = "變更路由"
'        Case "24"
'            strCmdName = "變更密碼"
'        Case "25"
'            strCmdName = "確認使用人"
'        Case "26"
'            strCmdName = "申請固定IP"
'        Case "27"
'            strCmdName = "取消申請固定IP"
'        Case "28"
'            strCmdName = "重置密碼"
'        Case "29"
'            strMsgCode = "16"
'            strCmdName = "FTTX同區移入"
'
'        Case 30
'            If blnUseFttx Then
'                strCmdName = "查詢FTTX資訊"
'            Else
'                strCmdName = "查詢CM資訊"
'            End If
'        Case 31
'            strCmdName = "查詢帳號資訊"
'        Case 32
'            strCmdName = "查詢FreePort"
'        Case 33
'            strCmdName = "Release Port"
'        Case 40
'            strCmdName = "設備入庫"
'        Case 41
'            strCmdName = "設備出庫"
'        Case Else
'            strCmdName = ""
'    End Select
'    SCM_Command = False
'    If strCmdName = "" Then strErrMsg = "找不到命令 !": GoTo 66
'    Set objCn = CreateObject("ADODB.Connection")
'
'    With objCn
'            .CursorLocation = 3
'            .Open strCn
'    End With
'    If objCn.State <= 0 Then
'        strErrMsg = "資料庫連線尚未開啟 !"
'        Exit Function
'    End If
'    strOwner = strOwner_Name
'    intCompCode = strComp_Code
'
'    On Error Resume Next
'    Dim varSysPara As Variant
'    Dim strQryPara As String
'    strQryPara = "SELECT INTEGRATESPM,CMTIMEOUT,CMOWNER FROM " & strOwner & "SO041" & _
'                                        " WHERE COMPCODE=" & intCompCode
'
'    varSysPara = objCn.Execute(strQryPara).GetRows(1, 1)
'
'    If Val(varSysPara(0, 0) & "") = 0 Then
'        strErrMsg = "[SO041] 系統業者參數檔 [啟動CMCP介接] 尚未開啟, 因此並未呼叫 Web Service 介接 !"
'        Exit Function
'    End If
'
'    lngCMtimeout = Val(varSysPara(1, 0) & "")
'
'    If lngCMtimeout = 0 Then lngCMtimeout = 30
'    strCMowner = varSysPara(2, 0) & ""
'    'strCMowner = GetOwner
'    If strCMowner = Empty Then
'        strCMowner = ""
'    Else
'        strCMowner = strCMowner & "."
'    End If
'    If Err.Number <> 0 Then
'        Err.Clear
'        strErrMsg = "[SO041] 系統業者參數檔 查無 [公司別] 為 " & intCompCode & " 的資料 , 請確認 !"
'        Exit Function
'    End If
'    strCmdSeqNo = Empty
'    If strCmdSeqNo = Empty Then str_CmdSeqNo = Get_Cmd_Seq_No(objCn)
'    strCmdSeqNo = str_CmdSeqNo
'    If Not Insert_SO314_Para(objCn, strErrMsg, strCmdSeqNo, str(intCompCode), strMsgCode, strCmdName, strOperatorId, _
'                                "W", strCustID, strCustName, strCustTel, strCustAddress, strDialAccount, strSchemaCode1, _
'                                strSchemaCode2, strUserIdNo, strUserName, strUserBirthday, strDeviceSNo1, strDeviceSNo2, _
'                                strDeviceType, strDeviceModelNo, strPassword, strSNO, strMediaBillNo, strProcessingDate, strReasonCode, strReasonName, , strParamData, strDynIpCount, strFixIPCount) Then GoTo 66
'    If strProcessingDate <> Empty Then
'        SCM_Command = True
'        GoTo 66
'    End If
'    If Not Pooling(objCn, strCustID, strErrMsg, strFaultReason, objPBG) Then GoTo 66
'    If Not rsResult Is Nothing Then Call GetScmData(objCn, rsResult)
'    SCM_Command = True
'66:
'    On Error Resume Next
'    objCn.Close
'    Set objCn = Nothing
'  Exit Function
'ChkErr:
'    ErrSub "clsSCMcommand", "SCM_Command"
'End Function
Private Function Insert_SO555_Para(ByRef objCn As Object, ByRef strRetErrMsg As String, _
                            ByVal strCMD As String, Optional strAction As String = "", _
                            Optional ByVal strCmdStatus As String, Optional ByVal strRequestTime As String = "", _
                            Optional ByVal strResponseTime As String = "", Optional ByVal strExchTime As String = "", _
                            Optional ByVal strRetMsg As String = "", _
                            Optional ByVal strSysMsg As String = "", Optional ByVal strServiceProviderID As String, _
                            Optional ByVal strAccountNum As String = "", Optional ByVal strAccoutName As String = "", _
                            Optional ByVal strAccountStatus As String = "", Optional ByVal strAddress As String = "", _
                            Optional ByVal strPostCode As String, Optional ByVal strAccountUID As String = Empty, _
                            Optional ByVal strLoginID As String = "", Optional ByVal strPin As String, _
                            Optional ByVal intDefaultUser As Integer, Optional ByVal intAccessPointUID As Integer, _
                            Optional ByVal strSmartCardID As String = "", Optional ByVal strMacAddress As String, _
                            Optional ByVal strSerialNumber As String = "", Optional ByVal strSTBstatus As String = "", _
                            Optional ByVal strDeviceType As String, Optional ByVal intDeviceEnabled As Integer, _
                            Optional ByVal strCASNo As String = "", Optional ByVal strProductID As String, _
                            Optional ByVal strProductName As String = "", Optional ByVal strUpdTime As String, Optional ByVal strUpdEn As String, _
                            Optional ByVal strProcessingDate As String = "", Optional ByVal strSNO As String = "", _
                            Optional ByVal strMediaBillNo As String = "", Optional ByVal strReasonCode As String = "", _
                            Optional ByVal strReasonName As String = "", Optional ByVal strOperatorId As String = "", _
                            Optional ByVal strSTBUId As String = "", Optional ByVal strCompCode As String = "", _
                            Optional ByVal strBillStartDate As String = "", Optional ByVal strBillEndDate As String = "", _
                            Optional ByVal strICCUID As String = "", Optional ByVal strCitemStr As String = "") As Boolean
  On Error GoTo chkErr
  Dim rs555 As Object
  Dim rs555A As Object
  Dim strPID As Variant
  Dim i As Long
  Set rs555 = CreateObject("ADODB.Recordset")
  Set rs555A = CreateObject("ADODB.Recordset")
  rs555.Open "Select * From " & strCMowner & "SO555 Where 1=0", objCn, adOpenKeyset, adLockOptimistic
  objCn.BeginTrans
    Call WriteRecordVodProcedure("2.2.2 開始新增SO555 > SEQNO = " & strCmdSeqNo)
    With rs555
        .AddNew
        .Fields("SeqNo").Value = strCmdSeqNo
        .Fields("CMD").Value = strCMD
        .Fields("Action").Value = strAction
        .Fields("CmdStatus").Value = strCmdStatus
        .Fields("ExchTime").Value = Format(strExchTime, "YYYY/MM/DD hh:mm:ss")       'Format(RightNow, "YYYY/MM/DD hh:mm:ss")
        .Fields("ServiceProviderID") = IIf(strServiceProviderID = "", Null, strServiceProviderID)
        .Fields("AccountNum") = IIf(strAccountNum = "", Null, strAccountNum)
        .Fields("AccoutName") = strAccoutName
        .Fields("AccountStatus") = strAccountStatus
        .Fields("Address") = strAddress
        .Fields("PostCode") = strPostCode
        If strAccountUID <> Empty Then
            .Fields("AccountUID") = strAccountUID
        End If
        .Fields("LoginID") = strLoginID
        .Fields("Pin").Value = strPin
        .Fields("DefaultUser").Value = intDefaultUser
        .Fields("AccessPointUID").Value = intAccessPointUID
        .Fields("SmartCardID").Value = strSmartCardID
        .Fields("MacAddress").Value = strMacAddress
        .Fields("SerialNumber").Value = strSerialNumber
        .Fields("STBstatus").Value = strSTBstatus
        .Fields("DeviceType").Value = strDeviceType
        .Fields("DeviceEnabled").Value = intDeviceEnabled
        .Fields("CASNo").Value = strCASNo
        .Fields("ProductID").Value = strProductID
        .Fields("ProductName").Value = strProductName
        .Fields("UpdTime").Value = GetDTString(strUpdTime)                        'Format(strUpdTime, "YYYY/MM/DD hh:mm:ss")
        .Fields("UpdEn").Value = strUpdEn
        .Fields("ResvTime").Value = IIf(strProcessingDate = "", Null, strProcessingDate)
        .Fields("SNO").Value = strSNO
        .Fields("MediaBillNo").Value = strMediaBillNo
        .Fields("ReasonCode").Value = IIf(strReasonCode = "", Null, strReasonCode)
        .Fields("ReasonName").Value = strReasonName
        .Fields("OperatorId").Value = strOperatorId
        .Fields("STBUID").Value = IIf(strSTBUId = "", Null, strSTBUId)
        .Fields("CompCode").Value = strCompCode
        If strBillStartDate <> "" Then
            .Fields("BillStartDate").Value = Format(strBillStartDate, "YYYY/MM/DD hh:mm:ss")
        End If
        If strBillEndDate <> "" Then
            .Fields("BillEndDate").Value = Format(strBillEndDate, "YYYY/MM/DD hh:mm:ss")
        End If
        '#5603 增加填入ICCUID By Kin 2010/03/29
        If strICCUID <> "" Then
            .Fields("ICCUID").Value = strICCUID
        End If
        .Update
    End With
    Call WriteRecordVodProcedure("2.2.2 新增SO555完成 ")
    rs555A.Open "Select * From " & strCMowner & "SO555A Where 1=0", objCn, adOpenKeyset, adLockOptimistic
    If strProductID = "" Then
        rs555A.AddNew
        rs555A("SeqNo") = strCmdSeqNo
        rs555A.Update
        strRetCmdSeqNo = strCmdSeqNo
    Else
        '#5708 SO555A增加存入傳入的CitemCode By Kin 2010/07/05
        strPID = Split(strProductID, ",")
        For i = LBound(strPID) To UBound(strPID)

            With rs555A
                .AddNew
                .Fields("SeqNo") = strCmdSeqNo
                If strPID(i) <> "" Then
                    
                    If Len(strCitemStr & "") > 0 Then
                        .Fields("ProductId") = GetRsValue("Select A.CodeNo From " & _
                                                        strCMowner & "CD024 A," & strCMowner & "CD019A B " & _
                                                        " Where A.ChannelId='" & strPID(i) & "'" & _
                                                        " AND A.CODENO=B.CODENO AND B.CITEMCODE IN(" & strCitemStr & ") " & _
                                                        " And A.VodType=2 And A.StopFlag<>1", objCn) & ""
    
                        .Fields("ProductName") = GetRsValue("Select A.Description From " & _
                                                strCMowner & "CD024 A, " & strCMowner & "CD019A B" & _
                                                " Where A.ChannelID='" & strPID(i) & "'" & _
                                                " AND A.CODENO=B.CODENO AND B.CITEMCODE IN(" & strCitemStr & ") " & _
                                                " And A.VodType=2 And A.StopFlag<>1", objCn) & ""
                    End If
                End If
                If Len(strCitemStr & "") > 0 Then
                    .Fields("CitemCode") = NoZero(GetChannelCitemCode(strPID(i), strCitemStr))
                End If
                .Update
            End With
        Next i
    End If
    Call WriteRecordVodProcedure("2.2.3 新增SO555A完成 ")
  objCn.CommitTrans
  Call WriteRecordVodProcedure("2.2.4 新增控制台完成 ")
  On Error Resume Next
  Call CloseRecordset(rs555)
  Insert_SO555_Para = True
  Exit Function
chkErr:
    objCn.RollbackTrans
    strRetErrMsg = "新增SO555資料錯誤!!"
    ErrSub "clsVodcommand", "Insert_SO555_Para"
End Function

Private Function Pooling(ByRef objCn As Object, _
                         ByVal strCustID As String, _
                         ByRef strErrMsg As String, _
                         ByRef strFaultReason As String, _
                         Optional ByRef objPGB As Object = Nothing) As Boolean
  On Error GoTo chkErr
    
    Dim lngTimer As Long
    Dim intTickCount As Integer
    Dim strPooling As String
    intTickCount = 0
    
    If Not objPGB Is Nothing Then objPGB.Max = lngCMtimeout
    Do Until intTickCount = lngCMtimeout
        If lngTimer = 0 Then lngTimer = GetTickCount
        If GetTickCount - lngTimer > 1000 Then
            lngTimer = GetTickCount
            If Not objPGB Is Nothing Then objPGB.Value = intTickCount
            strPooling = GetStateData(objCn, intCompCode)
            If UCase(strPooling) <> "W" And UCase(strPooling) <> "P" Then Exit Do
            intTickCount = intTickCount + 1
            DoEvents
        End If
    Loop
    If Not objPGB Is Nothing Then objPGB.Value = lngCMtimeout
    If UCase(strPooling) = "S" Then
        Pooling = True
    Else
        '#5469 線連逾時要將狀態改為E,SysMsg要填入連線逾時 By Kin 2010/01/07
        If (UCase(strPooling) = "W") Or UCase(strPooling) = "P" Then
            strErrMsg = "連線逾時!!"
            objCn.BeginTrans
            objCn.Execute "UPDATE " & strCMowner & "SO555 SET CmdStatus='E', " & _
                "SysMsg= '連線逾時' WHERE SEQNO=" & strCmdSeqNo
            objCn.CommitTrans
        Else
            Select Case UCase(strPooling)
                Case "E"
                    If Not GetErrMsg(objCn, intCompCode, strErrMsg, strFaultReason) Then
                        strErrMsg = "GateWay 系統執行有錯誤"
                    End If
                Case "T"
                    strErrMsg = "GateWay 系統執行Time Out"
                    If Not GetErrMsg(objCn, intCompCode, strErrMsg, strFaultReason) Then
                        strErrMsg = "GateWay 系統執行Time Out"
                    End If
                Case Else
                    If Not GetErrMsg(objCn, intCompCode, strErrMsg, strFaultReason) Then
                        strErrMsg = "處理完成!!，但有錯誤"
                    End If
            End Select
            
        End If
        Pooling = False
    End If
    
  Exit Function
chkErr:
    ErrSub "clsEMCcommand", "Pooling"
End Function

Private Sub Class_Terminate()
  On Error Resume Next
    lngCMtimeout = 0
    intCompCode = 0
End Sub

