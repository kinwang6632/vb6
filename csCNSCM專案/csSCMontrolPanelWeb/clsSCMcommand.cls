VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSCMcommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private lngCMtimeout As Long
'Private intCompCode As Integer

Public Function SCM_Command(ByVal strCn As String, ByVal strComp_Code As String, ByVal strOwner_Name As String, _
                            ByVal strMsgCode As String, Optional ByVal strOperatorId As String = Empty, Optional strCustID As String = Empty, _
                            Optional ByVal strCustName As String = Empty, Optional strCustTel As String = Empty, _
                            Optional ByVal strCustAddress As String = Empty, Optional ByVal strDialAccount As String = Empty, _
                            Optional ByVal strSchemaCode1 As String = Empty, Optional ByVal strSchemaCode2 As String = Empty, _
                            Optional ByVal strUserIdNo As String = Empty, Optional ByVal strUserName As String = Empty, _
                            Optional ByVal strUserBirthday As String = Empty, Optional ByVal strDeviceSNo1 As String = Empty, _
                            Optional ByVal strDeviceSNo2 As String = Empty, Optional ByVal strDeviceType As String = Empty, _
                            Optional ByVal strDeviceModelNo As String = Empty, Optional ByVal strPassword As String = Empty, _
                            Optional ByVal strSNO As String = Empty, Optional ByVal strMediaBillNo As String = Empty, _
                            Optional ByVal strProcessingDate As String = Empty, _
                            Optional ByVal strReasonCode As String = Empty, Optional ByVal strReasonName As String = Empty, _
                            Optional ByRef objPBG As Object = Nothing, Optional ByRef rsResult As Object = Nothing, _
                            Optional ByRef strResult As String = Empty, Optional ByRef strFaultReason As String = Empty, _
                            Optional ByRef strErrMsg As String = Empty, Optional ByVal strParamData As String = Empty, _
                            Optional ByVal strDynIpCount As String = Empty, Optional ByVal strFixIPCount As String = Empty) As Boolean
  On Error GoTo ChkErr
    Dim str_CmdSeqNo As String
    Dim strCmdName As String
    Dim objCn As Object
    strCmdName = Empty
    Select Case strMsgCode
        Case "10"
            If blnUseFttx Then
                strCmdName = "啟用PPPOE帳號"
            Else
                strCmdName = "建立CM服務帳號"
            End If
        Case "11"
            strCmdName = "開機"
        Case "12"
            strCmdName = "關機"
        Case "13"
            strCmdName = "停機"
        Case "14"
            strCmdName = "復機"
        '#4276 增加此共用指令 By Kin 2009/01/08
        Case "15"
            strCmdName = "終止服務帳號"
        Case "16"
            strCmdName = "設備更換"
        Case "17"
            strCmdName = "重置設備"
        Case "18"
            strCmdName = "預約FTTX裝機"
        Case "19"
            strCmdName = "FTTX裝機退單"
        Case "20"
            strCmdName = "變更基本資料"
        Case "21"
            strCmdName = "變更申請人"
        Case "22"
            strCmdName = "變更速率"
        Case "23"
            strCmdName = "變更路由"
        Case "24"
            strCmdName = "變更密碼"
        Case "25"
            strCmdName = "確認使用人"
        Case "26"
            strCmdName = "申請固定IP"
        Case "27"
            strCmdName = "取消申請固定IP"
        Case "28"
            strCmdName = "重置密碼"
        Case "29"
            strMsgCode = "16"
            strCmdName = "FTTX同區移入"
            
        Case 30
            If blnUseFttx Then
                strCmdName = "查詢FTTX資訊"
            Else
                strCmdName = "查詢CM資訊"
            End If
        Case 31
            strCmdName = "查詢帳號資訊"
        Case 32
            strCmdName = "查詢FreePort"
        '#4384 增加Release Port 命令 By Kin 2009/02/26
        Case 33
            strCmdName = "Release Port"
        Case 40
            strCmdName = "設備入庫"
        Case 41
            strCmdName = "設備出庫"
        Case Else
            strCmdName = ""
    End Select
    SCM_Command = False
    If strCmdName = "" Then strErrMsg = "找不到命令 !": GoTo 66
    Set objCn = CreateObject("ADODB.Connection")
    strCn = strCn & ";Persist Security Info=True"
    With objCn
            .CursorLocation = 3
            .Open strCn
    End With
    If objCn.State <= 0 Then
        strErrMsg = "資料庫連線尚未開啟 !"
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
        strErrMsg = "[SO041] 系統業者參數檔 [啟動CMCP介接] 尚未開啟, 因此並未呼叫 Web Service 介接 !"
        Exit Function
    End If
    
    lngCMtimeout = Val(varSysPara(1, 0) & "")

    If lngCMtimeout = 0 Then lngCMtimeout = 30
    strCMowner = varSysPara(2, 0) & ""
    'strCMowner = GetOwner
    If strCMowner = Empty Then
        strCMowner = ""
    Else
        strCMowner = strCMowner & "."
    End If
    If Err.Number <> 0 Then
        Err.Clear
        strErrMsg = "[SO041] 系統業者參數檔 查無 [公司別] 為 " & intCompCode & " 的資料 , 請確認 !"
        Exit Function
    End If
    strCmdSeqNo = Empty
    If strCmdSeqNo = Empty Then str_CmdSeqNo = Get_Cmd_Seq_No(objCn)
    strCmdSeqNo = str_CmdSeqNo
    Call WriteRecordVodProcedure("4.準備進入新增SO314")
    If Not Insert_SO314_Para(objCn, strErrMsg, strCmdSeqNo, str(intCompCode), strMsgCode, strCmdName, strOperatorId, _
                                "W", strCustID, strCustName, strCustTel, strCustAddress, strDialAccount, strSchemaCode1, _
                                strSchemaCode2, strUserIdNo, strUserName, strUserBirthday, strDeviceSNo1, strDeviceSNo2, _
                                strDeviceType, strDeviceModelNo, strPassword, strSNO, strMediaBillNo, strProcessingDate, strReasonCode, strReasonName, , strParamData, strDynIpCount, strFixIPCount) Then GoTo 66
    If strProcessingDate <> Empty Then
        SCM_Command = True
        GoTo 66
    End If
    If Not Pooling(objCn, strCustID, strErrMsg, strFaultReason, objPBG) Then GoTo 66
    If Not rsResult Is Nothing Then Call GetScmData(objCn, rsResult)
    SCM_Command = True
    
66:
    On Error Resume Next
    objCn.Close
    Set objCn = Nothing
  Exit Function
ChkErr:
    ErrSub "clsSCMcommand", "SCM_Command"
End Function

Private Function Insert_SO314_Para(ByRef objCn As Object, ByRef strErrMsg As String, ByVal strCmdSeqNo As String, _
                                                        ByVal strCompCode As String, ByVal strMsgCode As String, ByVal strMsgName As String, ByVal strOperatorId As String, _
                                                        ByVal strcmdStatus As String, Optional ByVal strCustID As String = "", Optional ByVal strCustName As String = "", _
                                                        Optional ByVal strCustTel As String = "", Optional ByVal strCustAddress As String = "", _
                                                        Optional ByVal strAccountID As String = "", Optional ByVal strSchemaCode1 As String = "", _
                                                        Optional ByVal strSchemaCode2 As String = "", Optional ByVal strUserIdNo As String = "", _
                                                        Optional ByVal strUserName As String = "", Optional ByVal strUserBirthday As String = "", _
                                                        Optional ByVal strDeviceSNo1 As String = "", Optional ByVal strDeviceSNo2 As String = "", _
                                                        Optional ByVal strDeviceType As String = "", Optional ByVal strDeviceModelNo As String = "", _
                                                        Optional ByVal strPassword As String = "", Optional ByVal strSNO As String = "", _
                                                        Optional ByVal strMediaBillNo As String = "", Optional ByVal strProcessingDate As String = Empty, _
                                                        Optional ByVal strReasonCode As String = Empty, Optional ByVal strReasonName As String = Empty, _
                                                        Optional ByVal strServiceType As String = Empty, Optional ByVal strParamData As String = Empty, _
                                                        Optional ByVal strDynIpCount As String = Empty, Optional ByVal strFixIPCount As String = Empty) As Boolean
                                                        
                                                    
  On Error GoTo ChkErr
    Dim lngRcdAft As Long
    Dim cmd As Object
    Dim rs314 As New ADODB.Recordset
    
    Set cmd = CreateObject("ADODB.Command")
    
    objCn.BeginTrans
'    If Not GetRS(rs314, "Select  From " & strCMowner & "SO314 Where 1=0", objCn, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'    rs314.AddNew
'    With rs314
'        .Fields("CmdSeqNo") = strCmdSeqNo
'        .Fields("CompCode") = strCompCode
'        .Fields("MsgCode") = strMsgCode
'    End With
'    rs314.Update
    
    '#4276 增加FTTX要存入的欄位 By Kin 2008/12/23
    With cmd
            Set .ActiveConnection = objCn
            .CommandType = 1
            .CommandText = "INSERT INTO " & strCMowner & "SO314 " & _
                                        "(CmdSeqNo,CompCode,MsgCode,MsgName,OperatorId,CmdStatus,CustID,CustName" & _
                                        ",CustTel,CustAddress,AccountId,SchemaCode1,SchemaCode2,UserIdNo,UserName" & _
                                        ",UserBirthday ,DeviceSNo1,DeviceSNo2,DeviceType,DeviceModelNo,UpdTime,AccountPassword,SNO,MEDIABILLNO," & _
                                        "ProcessingDate,ReasonCode,ReasonName,ServiceType,ParamData,DynIpCount,FixIPCount)" & _
                                        " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,TO_DATE(?,'YYYY/MM/DD'),?,?,?,?,TO_DATE(?,'YYYY/MM/DD HH24:MI:SS'),?,?,?,TO_DATE(?,'YYYY/MM/DD HH24:MI:SS'),?,?,?,?,?,?)"




            On Error GoTo ChkErr

           

             .Execute lngRcdAft, Array(strCmdSeqNo, strCompCode, strMsgCode, strMsgName, _
                                        strOperatorId, strcmdStatus, IIf(strCustID = Empty, Null, strCustID), IIf(strCustName = Empty, Null, strCustName), _
                                        IIf(strCustTel = Empty, Null, strCustTel), IIf(strCustAddress = Empty, Null, strCustAddress), IIf(strAccountID = Empty, Null, strAccountID), IIf(strSchemaCode1 = Empty, Null, strSchemaCode1), _
                                        strSchemaCode2, strUserIdNo, strUserName, IIf(strUserBirthday = "", Null, Format(strUserBirthday, "YYYY/MM/DD")), _
                                        strDeviceSNo1, strDeviceSNo2, strDeviceType, strDeviceModelNo, Format(RightNow, "YYYY/MM/DD hh:mm:ss"), strPassword, strSNO, strMediaBillNo, strProcessingDate, strReasonCode, strReasonName, _
                                        IIf(blnUseFttx, "FTTX", Null), strParamData, strDynIpCount, strFixIPCount)


    End With
    Call WriteRecordVodProcedure("5.新增SO314完成 > SEQNO = " & strCmdSeqNo)
    objCn.CommitTrans
    Insert_SO314_Para = True
    On Error Resume Next
    Set cmd = Nothing
    Call CloseRecordset(rs314)
  Exit Function
ChkErr:
    objCn.RollbackTrans
    strErrMsg = "新增SO314資料錯誤!!"
    ErrSub "clsSCMcommand", "Insert_2_SO314_Para"
End Function

Private Function Pooling(ByRef objCn As Object, _
                         ByVal strCustID As String, _
                         ByRef strErrMsg As String, _
                         ByRef strFaultReason As String, _
                         Optional ByRef objPGB As Object = Nothing) As Boolean
  On Error GoTo ChkErr
    
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
            If blnShowMsg Then
                DoEvents
            End If
        End If
    Loop
    If Not objPGB Is Nothing Then objPGB.Value = lngCMtimeout
    If UCase(strPooling) = "C" Then
        Pooling = True
    Else
        If (UCase(strPooling) = "W") Or UCase(strPooling) = "P" Then
            strErrMsg = "連線逾時!!"
        Else
            If Not GetErrMsg(objCn, intCompCode, strErrMsg, strFaultReason) Then
                strErrMsg = "處理完成!!，但有錯誤"
            End If
        End If
        Pooling = False
    End If
    Call WriteRecordVodProcedure("6.Pooling完成")
  Exit Function
ChkErr:
    ErrSub "clsEMCcommand", "Pooling"
End Function

Private Sub Class_Terminate()
  On Error Resume Next
    lngCMtimeout = 0
    intCompCode = 0
End Sub

