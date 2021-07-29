Attribute VB_Name = "mod_CallSF"
'    If Call_IEL_00(strCompCode, strCmdSno, strStatus, lngFailureCode, lngReturn) Then
'        Debug.Print "00. 查詢SPM指令執行狀況 Message Get Command Status ( GET_COMMAND_STATUS ) 送出完成 !"
'        Qry_Result = True
'        If Not blnNoShwMsg Then
'            If UCase(strStatus) <> "N" Or blnTimeOut Then ShowMsg
'        End If
'    Else
'        Qry_Result = False
'        Debug.Print "00. 查詢SPM指令執行狀況 Message Get Command Status ( GET_COMMAND_STATUS ) 送出失敗 !"
'        If blnShowMsg And Not blnNoShwMsg Then
'            If UCase(strStatus) <> "N" Or blnTimeOut Then ShowMsg
'        End If
'    End If

'Private Sub ShowResult()
'  On Error GoTo ChkErr
'    Dim mflds As New GiGridFlds
'    Dim intLoop As Integer
'    frmSO1623A.ggrQueryInfo.Blank = True
'    With rs309
'            If .RecordCount = 0 Then Exit Sub
'            .MoveFirst
'            For intLoop = 0 To .Fields.Count - 1
'                SetGrdFld mflds, .Fields(intLoop).Name
'                DoEvents
'            Next
'    End With
'    With frmSO1623A.ggrQueryInfo
'            .AllFields = mflds
'            Set .Recordset = rs309
'            .Refresh
'            .Blank = False
'    End With
'  Exit Sub
'ChkErr:
'    ErrSub FrmName, "ShowResult"
'End Sub

'Public Function Insert2Table()
'    Dim objCmd As Object
'    ' Dim gcnGi As Object
'    Dim lngRcdAft As Long
'
'    Set objCmd = CreateObject("ADODB.Command")
'    ' gcnGi.BeginTrans
'
'    With objCmd
'            Set .ActiveConnection = gcnGi
'            .CommandType = 1
'            .CommandText = "INSERT INTO " & GetOwner & "SO309" & _
'                                        " (CMDSEQNO,PRIORITY,SOURCE,MSGCODE,MSGNAME,COMPCODE,CUSTID,MODEMMAC,NEWMODEMMAC" & _
'                                        ",MODELNAME,NEWMODELNAME,CMBAUDRATE,NEWCMBAUDRATE,DYNIPCOUNT,NEWDYNIPCOUNT,FIXIPCOUNT,NEWFIXIPCOUNT,MEDIAGATEWAY" & _
'                                        ",CPNUMBER,TEL,HFCNODE,NEWHFCNODE,CPEMAC,ADDCPEMAC,DELCPEMAC,CPESTATIP,ADDCPESTATIP" & _
'                                        ",DELCPESTATIP,CPEIPRULE,IPADDRESS,NEWIPADDRESS,CMIPRULE,MTAMAC,NEWMTAMAC,COMMANDSTATUS,RETCODE" & _
'                                        ",RETMSG,CLIFAILURECODE,EXECENTRY,EXECTIME,GATEEXECTIME,GATEFINSHTIME,GATEROLLBACKTIME,HOSTNAME,SNO" & _
'                                        ",ORDERNO,RESVTIME,CMDRESEND,QCMIPADDRESS,QCPEIPADDRESS,QDESCRIPTION,QUPTIME,QSERIAL,QSTATUS" & _
'                                        ",QDOWNSTREAMFREQUENCY,QDOWNSTREAMPOWER,QERRORLESS,QCORRECTABLE,QUNCORRECTABLE,QDOWNSTREAMSNRATIO,QMICROREFLECTIONS,QDOWNSTREAMMODULATION,QUPSTREAMFREQUENCY" & _
'                                        ",QUPSTREAMPOWER,QTIMINGOFFSET,QRESETS,QLOSTSYNC,QINVALIDMAPS,QINVALIDUCDS,QINVALIDRANGING,QINVALIDREGISTRATIO,QT1TIMEOUTS" & _
'                                        ",QT2TIMEOUTS,QT3TIMEOUTS,QT4TIMEOUTS,QRANGINGABORTS,QQOSSTATUS,QQOSPRIORITY,QQOSGUBANDWIDTH,QQOSMUBANDWIDTH,QQOSMDBANDWIDTH" & _
'                                        ",QQOSMTBURST,QQOSBPIENABLED)" & _
'                                        " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
'             .Execute lngRcdAft, Array("[CMDSEQNO]", "[PRIORITY]", "[SOURCE]", "[MSGCODE]", "[MSGNAME]", "[COMPCODE]", "[CUSTID]", "[MODEMMAC]", "[NEWMODEMMAC]" _
'                                                        , "[MODELNAME]", "[NEWMODELNAME]", "[CMBAUDRATE]", "[NEWCMBAUDRATE]", "[DYNIPCOUNT]", "[NEWDYNIPCOUNT]", "[FIXIPCOUNT]", "[NEWFIXIPCOUNT]", "[MEDIAGATEWAY]" _
'                                                        , "[CPNUMBER]", "[TEL]", "[HFCNODE]", "[NEWHFCNODE]", "[CPEMAC]", "[ADDCPEMAC]", "[DELCPEMAC]", "[CPESTATIP]", "[ADDCPESTATIP]" _
'                                                        , "[DELCPESTATIP]", "[CPEIPRULE]", "[IPADDRESS]", "[NEWIPADDRESS]", "[CMIPRULE]", "[MTAMAC]", "[NEWMTAMAC]", "[COMMANDSTATUS]", "[RETCODE]" _
'                                                        , "[RETMSG]", "[CLIFAILURECODE]", "[EXECENTRY]", "[EXECTIME]", "[GATEEXECTIME]", "[GATEFINSHTIME]", "[GATEROLLBACKTIME]", "[HOSTNAME]", "[SNO]" _
'                                                        , "[ORDERNO]", "[RESVTIME]", "[CMDRESEND]", "[QCMIPADDRESS]", "[QCPEIPADDRESS]", "[QDESCRIPTION]", "[QUPTIME]", "[QSERIAL]", "[QSTATUS]" _
'                                                        , "[QDOWNSTREAMFREQUENCY]", "[QDOWNSTREAMPOWER]", "[QERRORLESS]", "[QCORRECTABLE]", "[QUNCORRECTABLE]", "[QDOWNSTREAMSNRATIO]", "[QMICROREFLECTIONS]", "[QDOWNSTREAMMODULATION]", "[QUPSTREAMFREQUENCY]" _
'                                                        , "[QUPSTREAMPOWER]", "[QTIMINGOFFSET]", "[QRESETS]", "[QLOSTSYNC]", "[QINVALIDMAPS]", "[QINVALIDUCDS]", "[QINVALIDRANGING]", "[QINVALIDREGISTRATIO]", "[QT1TIMEOUTS]" _
'                                                        , "[QT2TIMEOUTS]", "[QT3TIMEOUTS]", "[QT4TIMEOUTS]", "[QRANGINGABORTS]", "[QQOSSTATUS]", "[QQOSPRIORITY]", "[QQOSGUBANDWIDTH]", "[QQOSMUBANDWIDTH]", "[QQOSMDBANDWIDTH]" _
'                                                        , "[QQOSMTBURST]", "[QQOSBPIENABLED]")
'    End With
'
'    ' gcnGi.CommitTrans
'
'    On Error Resume Next
'    Set objCmd = Nothing
'End Function
'
'Private Sub Chk_Resend(str_Cmd_Sno As String)
'  On Error GoTo ChkErr
'    If str_Cmd_Sno <> Empty Then
'        gcnGi.Execute "UPDATE " & GetOwner & "SO309" & _
'                                " SET CMDRESEND=" & strCmdSno & _
'                                " WHERE CMDSEQNO='" & str_Cmd_Sno & "'"
'    End If
'  Exit Sub
'ChkErr:
'    ErrSub "mod_ControlPanel", "Chk_Resend"
'End Sub
'
'
'

'Private Sub SetGrdFld(mflds As GiGridFlds, strFldName As String)
'  On Error GoTo ChkErr
'    With mflds
'            Select Case UCase(strFldName)
'                        Case "CMDSEQNO": .Add "CMDSEQNO", , , , , "指令流水序號       ", vbLeftJustify
'                        Case "MSGCODE": .Add "MSGCODE", , , , , "執行代碼", vbRightJustify
'                        Case "MSGNAME": .Add "MSGNAME", , , , , "執行名稱                 ", vbLeftJustify
'                        Case "SERVICEPROVIDER": .Add "SERVICEPROVIDER", , , , , "公司別", vbRightJustify
'                        Case "ACCOUNTNUM": .Add "ACCOUNTNUM", , , , , "客戶編號", vbRightJustify
'                        Case "MODEMMAC": .Add "MODEMMAC", , , , , "CM MAC Address", vbLeftJustify
'                        Case "MODEMMACNEW": .Add "MODEMMACNEW", , , , , "換新CM MAC       ", vbLeftJustify
'                        Case "HFCNODE": .Add "HFCNODE", , , , , "網路編號(光投落點)", vbLeftJustify
'                        Case "INTERNETTYPENAME": .Add "INTERNETTYPENAME", , , , , "CM速率  ", vbLeftJustify
'                        Case "EQUIPTYPE": .Add "EQUIPTYPE", , , , , "型號MODEL ", vbLeftJustify
'                        Case "NUMDYNAIP": .Add "NUMDYNAIP", , , , , "動態IP數目", vbRightJustify
'                        Case "NUMSTATIP": .Add "NUMSTATIP", , , , , "固定IP數目", vbRightJustify
'                        Case "CPEMAC": .Add "CPEMAC", , , , , "網卡MAC         ", vbLeftJustify
'                        Case "CPEMACNEW": .Add "CPEMACNEW", , , , , "新網卡MAC    ", vbLeftJustify
'                        Case "EQUIPSTATUS": .Add "EQUIPSTATUS", , , , , "SPM設定狀態", vbLeftJustify
'                        Case "IPADDRESS": .Add "IPADDRESS", , , , , "EMTA IP Address", vbLeftJustify
'                        Case "VOIPTYPE": .Add "VOIPTYPE", , , , , "服務別", vbLeftJustify
'                        Case "MTAMAC": .Add "MTAMAC", , , , , "MTA MAC             ", vbLeftJustify
'                        Case "MTAMACNEW": .Add "MTAMACNEW", , , , , "MTA MAC NEW ", vbLeftJustify
'                        Case "MEDIAGATEWAY": .Add "MEDIAGATEWAY", , , , , "EMTA Media Gateway", vbLeftJustify
'                        Case "RETCODE": .Add "RETCODE", , , , , "回應代碼攔", vbRightJustify
'                        Case "RETMSG": .Add "RETMSG", , , , , "回應訊息攔                          ", vbLeftJustify
'                        Case "STATUSCODE": .Add "STATUSCODE", , , , , "SPM執行結果", vbLeftJustify
'                        Case "FAILURECODE": .Add "FAILURECODE", , , , , "SPM執行結果名稱  ", vbRightJustify
'                        Case "CLIENTCLASSNAME": .Add "CLIENTCLASSNAME", , , , , "設備狀態回應內容", vbLeftJustify
'                        Case "SOURCE": .Add "SOURCE", , , , , "來源             ", vbLeftJustify
'                        Case "EXECENTRY": .Add "EXECENTRY", , , , , "執行者           ", vbLeftJustify
'                        Case "EXECDATE": .Add "EXECDATE", , , , , "執行日期         ", vbLeftJustify
'                        Case "HOSTNAME": .Add "HOSTNAME", , , , , "電腦名稱           ", vbLeftJustify
'                        Case "SNO": .Add "SNO", , , , , "工單單號              ", vbLeftJustify
'                        Case "RESVTIME": .Add "RESVTIME", giControlTypeDate, , , , "預約日期時間 ", vbLeftJustify
'                        Case "CMDRESEND": .Add "CMDRESEND", , , , , "重送指令新的指令流水號", vbLeftJustify
'                        Case "CMDFLAG": .Add "CMDFLAG", , , , , "是否已傳送SPM", vbRightJustify
'                        Case "Q_RETURN_CODE": .Add "Q_RETURN_CODE", , , , , "查詢指令回傳值", vbLeftJustify
'                        Case "Q_STATUS_CODE": .Add "Q_STATUS_CODE", , , , , "查詢指令執行結果", vbLeftJustify
'                        Case "Q_FAILURE_CODE": .Add "Q_FAILURE_CODE", , , , , "執行結果的指令交易錯誤訊息代碼", vbLeftJustify
'                        Case "Q_IPADDRESS": .Add "Q_IPADDRESS", , , , , "SPM 查詢指令結果IPADDRESS", vbLeftJustify
'                        Case "Q_MACADDRESS": .Add "Q_MACADDRESS", , , , , "SPM 查詢指令結果MACADDRESS", vbLeftJustify
'                        Case "Q_LEASETIME": .Add "Q_LEASETIME", , , , , "SPM 查詢指令結果LEASETIME", vbLeftJustify
'                        Case "Q_INTERFACE": .Add "Q_INTERFACE", , , , , "SPM 查詢指令結果INTERFACE", vbLeftJustify
'                        Case "Q_ADDRESSSTATE": .Add "Q_ADDRESSSTATE", , , , , "SPM 查詢指令結果ADDRESSSTATE", vbLeftJustify
'                        Case "Q_IPLEASESTARTTIME": .Add "Q_IPLEASESTARTTIME", , , , , "SPM 查詢指令結果IPLEASESTARTTIME", vbLeftJustify
'                        Case "Q_IPLEASEINITSTARTTIME": .Add "Q_IPLEASEINITSTARTTIME", , , , , "SPM 查詢指令結果IPLEASEINITSTARTTIME", vbLeftJustify
'                        Case "Q_IPLEASEPROTOCOL": .Add "Q_IPLEASEPROTOCOL", , , , , "SPM 查詢指令結果IPLEASEPROTOCOL", vbLeftJustify
'                        Case "Q_IPLEASEREMOTEID": .Add "Q_IPLEASEREMOTEID", , , , , "SPM 查詢指令結果IPLEASEREMOTEID", vbLeftJustify
'                        Case "Q_SUBNETMASK": .Add "Q_SUBNETMASK", , , , , "SPM 查詢指令結果SUBNETMASK", vbLeftJustify
'                        Case "Q_DOMAINSERVER": .Add "Q_DOMAINSERVER", , , , , "SPM 查詢指令結果DOMAINSERVER", vbLeftJustify
'                        Case "Q_HOSTNAME": .Add "Q_HOSTNAME", , , , , "SPM 查詢指令結果HOSTNAME", vbLeftJustify
'                        Case "Q_TFTPSERVERNAME": .Add "Q_TFTPSERVERNAME", , , , , "SPM 查詢指令結果TFTPSERVERNAME", vbLeftJustify
'                        Case "Q_BOOTFILENAME": .Add "Q_BOOTFILENAME", , , , , "SPM 查詢指令結果BOOTFILENAME", vbLeftJustify
'                        Case "Q_GATEWAYS": .Add "Q_GATEWAYS", , , , , "SPM 查詢指令結果GATEWAYS", vbLeftJustify
'        End Select
'    End With
'  Exit Sub
'ChkErr:
'    ErrSub FrmName, "SetGrdFld"
'End Sub
'
'Private Function GetRS309(Optional strFlds As String = "", Optional strKey As String = "") As Boolean
'  On Error GoTo ChkErr
'    Dim strWhere As String
'    With rs309
'            If .State > 0 Then .Close
'            .CursorLocation = adUseClient
'            .CacheSize = 3
'             If strKey <> Empty Then
'                strWhere = " WHERE CMDSEQNO='" & strKey & "'"
'             Else
'                strWhere = " WHERE 0=1"
'             End If
'             If strFlds = Empty Then
'                .Open "SELECT * FROM " & GetOwner & "SO309" & strWhere, gcnGi, adOpenStatic, adLockOptimistic
'             Else
'                .Open "SELECT " & strFlds & " FROM " & GetOwner & "SO309" & strWhere, gcnGi, adOpenStatic, adLockOptimistic
'             End If
'             GetRS309 = (.State > 0)
'    End With
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "GetRS309"
'End Function
'
'Private Sub Update309()
'  On Error GoTo ChkErr
'    With rs309
'            If lngReturn < 0 Then lngReturn = 1
'            .Fields("RetCode") = lngReturn
'            .Fields("StatusCode") = strStatus
'            .Fields("FailureCode") = lngFailureCode
'             If InStr(1, strErrMsg, vbCrLf) Then
'                .Fields("RetMsg") = Split(strErrMsg, vbCrLf)(0)
'             Else
'                .Fields("RetMsg") = strErrMsg
'             End If
'            .Fields("Q_Status_Code") = strStatus
'            .Fields("Q_Failure_Code") = lngFailureCode
'    End With
'  Exit Sub
'ChkErr:
'    ErrSub FrmName, "Update309"
'End Sub

'Public Function Cmd_01(str_Dev_Status As String, _
'                                        str_CM_MAC As String, _
'                                        str_MTA_MAC As String, _
'                                        str_Model As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean ' 1. CM入庫 [01]
'
'  On Error GoTo ChkErr
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM," & _
'                            "MODEMMAC,MTAMAC,EQUIPSTATUS,EQUIPTYPE,RETCODE,RETMSG," & _
'                            "STATUSCODE,FAILURECODE,Q_STATUS_CODE,Q_FAILURE_CODE,CMDFLAG," & _
'                            "SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strCmdName = "01. CM 入庫"
'    strCompCode = GetSvcPvd
'    strCMmac = str_CM_MAC
'    strMTAMAC = str_MTA_MAC
'    strModel = str_Model
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "01", "EQUIPMENT_CREATE"
'
'    With rs309
'            If .State > 0 Then
'                .Fields("MTAMAC") = strMTAMAC
'                .Fields("EquipStatus") = str_Dev_Status
'                .Fields("EquipType") = strModel
'                .Update
'            End If
'    End With
'
'    If Call_IEL_01(strCompCode, strCmdSno, "", "", strCMmac, strMTAMAC, str_Dev_Status, strModel, lngMsgID, lngReturn) Then
'        Debug.Print "01. CM 入庫 Add modem to inventory ( EQUIPMENT_CREATE ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            ' OK !
'        End If
'    Else
'        ShowMsg
'    End If
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_01"
'End Function
'
'Private Sub Insert309(strSno As String, strMsgCode As String, strMsgName As String)
'  On Error GoTo ChkErr
'    If GetRS309(str309SQL) Then
'        With rs309
'            .AddNew
'            .Fields("ServiceProvider") = strCompCode
'            .Fields("CMDSeqNo") = strCmdSno
'            .Fields("ModemMAC") = strCMmac
'            If lngReturn > 0 Then
'                .Fields("RetCode") = lngReturn
'            Else
'                .Fields("RetMsg") = lngReturn
'            End If
'            .Fields("MsgCode") = strMsgCode
'            .Fields("MsgName") = strMsgName
'            .Fields("AccountNum") = frmSO1623A.ginCustId.Text
'            .Fields("Source") = "Cablesoft"
'            .Fields("ExecEntry") = garyGi(1)
'            .Fields("ExecDate") = GetDTString(RightNow)
'            .Fields("HostName") = Get_Computer_Name
'            .Fields("SNO") = strSno
'            .Fields("CMDFLAG") = 1
'        End With
'    End If
'  Exit Sub
'ChkErr:
'    ErrSub FrmName, "Insert309"
'End Sub
'
'' 丙、 02 服務設定由設備資料選定後按照服務設定所需參數取出SO004的相關資料與畫面設定值由設備預設帶出有
''   (固定ip無cpe mac 時則按照固定ip 數目傳送'NULL#NULL'或'MA:C:MM:MM:MM#NULL 以null補足固定ip 對應的mac)，
''   若有異動已畫面異動直傳送指令，指令執行結果(第二道查詢結果)若成功需要回填SO004的開機時間
'Public Function Cmd_02(str_CM_MAC As String, _
'                                        str_Cust_ID As String, _
'                                        str_HFC_Node As String, _
'                                        str_CM_Baud_Rate As String, _
'                                        str_VOIP_Type As String, _
'                                        str_Media_Gateway As String, _
'                                        str_IP_Address As String, _
'                                        str_Dyna_IP_Count As String, _
'                                        str_Fix_IP_Count As String, _
'                                        str_CPE_MAC As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean                                             ' 3. 服務設定 [02]
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123000300001','','','00:08:0E:AC:A8:9C','6666','Z28','2M','','','','1','1','A0:00:00:00:00:01',:vMsgID
'
''    Dim str_Cust_ID As String
''    Dim str_HFC_Node As String ' HFC Node : SO014.CircuitNo 或 NodeNo 由SO041.EMTAIPTYPE 決定 ( 0=網路編號 , 1=Node )
''    Dim str_CM_Baud_Rate As String ' CM BaudRate : SO004.CMBaudRate
''    Dim str_VOIP_Type As String ' VOIP Type : 有CP 服務時 需傳送 CD043.VoiceType ( N )
''    Dim str_Media_Gateway As String ' Media Gateway : 有CP 服務時 SO049.Gateway CP 服務由門號決定 ( Gateway ) ( N )
''    Dim str_IP_Address As String ' IP Address : 有CP 服務時  SO004.IPAddress 當品名參考號屬於EMTA 5 ( N )
''    Dim str_Dyna_IP_Count As String ' Number Dynamic IP : SO004.DynIPCount 需考慮增加或減少，以最後總數傳送，不應該 < 0 ( N )
''    Dim str_Fix_IP_Count As String ' Number Static IP : SO004.FixIPCount 需考慮增加或減少，以最後總數傳送，不應該 < 0 ( N )
''    Dim str_CPE_MAC As String ' CPE MAC Address : SO004C.CPEMAC ( N )
'
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM," & _
'                            "MODEMMAC,HFCNODE,INTERNETTYPENAME,VOIPTYPE,MEDIAGATEWAY," & _
'                            "IPADDRESS,NUMDYNAIP,NUMSTATIP,CPEMAC,RETCODE,RETMSG," & _
'                            "STATUSCODE,FAILURECODE,Q_STATUS_CODE,Q_FAILURE_CODE,CMDFLAG," & _
'                            "SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strCmdName = "02. 裝機"
''    strCmdSno = Get_Cmd_Sno
'    strCompCode = GetSvcPvd
'    strCMmac = str_CM_MAC
'
''    With frmSO1623A
''            str_Cust_ID = .ginCustId.Text
''            str_HFC_Node = IIf(int_EMTA_IP_Type = 0, .strCircuitNo, .strNodeNo)
''            str_CM_Baud_Rate = rsData("CMBaudRate") & ""
''            str_VOIP_Type = IIf(.IsCMCP, Get_VOIP_Type(rsData("ModelCode") & ""), "")
''            str_Media_Gateway = rsData("Gateway") & ""
''            str_IP_Address = "" ' rsData("IPAddress") & ""
''            str_Dyna_IP_Count = .txtDynaIP
''            str_Fix_IP_Count = .txtFixIP
''            str_CPE_MAC = Get_CPE_Mac(Val(str_Fix_IP_Count))
''    End With
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "02", "EQUIPMENT_ASSIGN"
'
'    With rs309
'            If .State > 0 Then
'                .Fields("HFCNODE") = str_HFC_Node
'                .Fields("INTERNETTYPENAME") = str_CM_Baud_Rate
'                .Fields("VOIPTYPE") = str_VOIP_Type
'                .Fields("MEDIAGATEWAY") = str_Media_Gateway
'                .Fields("IPADDRESS") = str_IP_Address
'                .Fields("NUMDYNAIP") = Val(str_Dyna_IP_Count)
'                .Fields("NUMSTATIP") = Val(str_Fix_IP_Count)
'                .Fields("CPEMAC") = str_CPE_MAC
'                .Update
'            End If
'    End With
'
'    If Call_IEL_02(strCompCode, strCmdSno, "", "", strCMmac, str_Cust_ID, str_HFC_Node, _
'                            str_CM_Baud_Rate, str_VOIP_Type, str_Media_Gateway, str_IP_Address, _
'                            str_Dyna_IP_Count, str_Fix_IP_Count, str_CPE_MAC, lngMsgID, lngReturn) Then
'        Debug.Print "02. 裝機 Installation ( EQUIPMENT_ASSIGN ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            ' OK !
'            If UCase(strStatus) = "S" Then
'                gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
'                                        " SET CMOPENDATE=" & O2Date(RightNow) & _
'                                        ",UPDEN='CABLESOFT'" & _
'                                        ",UPDTIME='" & GetDTString(RightNow) & "'" & _
'                                        " WHERE FACISNO='" & Replace(str_CM_MAC, ":", "") & "'"
'            End If
''            With rsData
''                .Fields("CMOpenDate") = Format(RightNow, strYMDHMS)
''                .Fields("UpdEn") = "CABLESOFT"
''                .Fields("UpdTime") = GetDTString(RightNow)
''                .Update
''            End With
'        End If
'    Else
'        ShowMsg
'    End If
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_02"
'End Function
'
'' 丁、 03 服務刪除由設備資料選定後按照服務設定所需參數取出SO004的相關資料，指令執行結果(第二道查詢結果)
''   若成功需要回填SO004的關機時間
'Public Function Cmd_03(str_CM_MAC As String, _
'                                        str_Cust_ID As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean ' 4. 服務刪除 [03]
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123000400001','','','00:0E:5C:60:61:7E','INACTIVE','8888',:vMsgID
'
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM," & _
'                            "MODEMMAC,EQUIPSTATUS,RETCODE,RETMSG," & _
'                            "STATUSCODE,FAILURECODE,Q_STATUS_CODE,Q_FAILURE_CODE,CMDFLAG," & _
'                            "SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strCmdName = "03 . 裝機退件 / 拆機 (分機) 取回設備"
'    'strCmdSno = Get_Cmd_Sno
'    strCompCode = GetSvcPvd
'    strCMmac = str_CM_MAC 'Mask(rsData("FaciSno") & "")
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "03", "EQUIPMENT_DEASSIGN"
'
'    With rs309
'            If .State > 0 Then
'                .Fields("EquipStatus") = "INACTIVE"
'                .Update
'            End If
'    End With
'
'    If Call_IEL_03(strCompCode, strCmdSno, "", "", strCMmac, str_Cust_ID, "INACTIVE", lngMsgID, lngReturn) Then
'        Debug.Print "03 . 裝機退件 / 拆機 (分機) 取回設備  Install Reject / Disconnect ( EQUIPMENT_DEASSIGN ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            ' OK !
'            If UCase(strStatus) = "S" Then
'                gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
'                                        " SET CMCLOSEDATE=" & O2Date(RightNow) & _
'                                        ",UPDEN='CABLESOFT'" & _
'                                        ",UPDTIME='" & GetDTString(RightNow) & "'" & _
'                                        " WHERE FACISNO='" & Replace(str_CM_MAC, ":", "") & "'"
'            End If
''            With rsData
''                .Fields("CMCloseDate") = Format(RightNow, strYMDHMS)
''                .Fields("UpdEn") = "CABLESOFT"
''                .Fields("UpdTime") = GetDTString(RightNow)
''                .Update
''            End With
'        End If
'    Else
'        ShowMsg
'    End If
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_03"
'End Function
'
'' 甲、 04 軟拆 由設備資料選定後按照服務設定所需參數取出SO004的相關資料，指令執行結果(第二道查詢結果)
''   若成功需要回填SO004的關機時間
'Public Function Cmd_04(str_CM_MAC As String, _
'                                        str_Cust_ID As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean ' 1. 軟拆 [04]
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123000700001','','','00:08:0E:AC:A8:9C','6666',:vMsgID
'
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM," & _
'                            "MODEMMAC,RETCODE,RETMSG,STATUSCODE,FAILURECODE," & _
'                            "Q_STATUS_CODE,Q_FAILURE_CODE,CMDFLAG," & _
'                            "SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strCmdName = "04 . 軟拆"
''    strCmdSno = Get_Cmd_Sno
'    strCompCode = GetSvcPvd
'    strCMmac = str_CM_MAC ' strCMMAC = Mask(rsData("FaciSno") & "")
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "04", "EQUIPMENT_LOCK"
'
'    With rs309
'            If .State > 0 Then .Update
'    End With
'
'    If Call_IEL_04(strCompCode, strCmdSno, "", "", strCMmac, str_Cust_ID, lngMsgID, lngReturn) Then
'        Debug.Print "04 . 軟拆 Soft Disconnect ( EQUIPMENT_LOCK ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            ' OK !
'            If UCase(strStatus) = "S" Then
'                gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
'                                        " SET CMCLOSEDATE=" & O2Date(RightNow) & _
'                                        ",UPDEN='CABLESOFT'" & _
'                                        ",UPDTIME='" & GetDTString(RightNow) & "'" & _
'                                        " WHERE FACISNO='" & Replace(str_CM_MAC, ":", "") & "'"
'            End If
''            With rsData
''                .Fields("CMCloseDate") = Format(RightNow, strYMDHMS)
''                .Fields("UpdEn") = "CABLESOFT"
''                .Fields("UpdTime") = GetDTString(RightNow)
''                .Update
''            End With
'        End If
'    Else
'        ShowMsg
'    End If
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_04"
'End Function
'
'' 乙、 05 軟復由設備資料選定後按照服務設定所需參數取出SO004的相關資料，指令執行結果(第二道查詢結果)
''   若成功需要回填 SO004 的開機時間
'Public Function Cmd_05(str_CM_MAC As String, _
'                                        str_Cust_ID As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean ' 2. 軟復 [05]
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123000800002','','','00:50:BA:2F:00:79','8888',:vMsgID
'
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM," & _
'                            "MODEMMAC,RETCODE,RETMSG,STATUSCODE,FAILURECODE," & _
'                            "Q_STATUS_CODE,Q_FAILURE_CODE,CMDFLAG," & _
'                            "SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strCmdName = "05 . 軟復"
''    strCmdSno = Get_Cmd_Sno
'    strCompCode = GetSvcPvd
'    strCMmac = str_CM_MAC ' strCMMAC = Mask(rsData("FaciSno") & "")
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "05", "EQUIPMENT_UNLOCK"
'
'    With rs309
'            If .State > 0 Then .Update
'    End With
'
'    If Call_IEL_05(strCompCode, strCmdSno, "", "", strCMmac, str_Cust_ID, lngMsgID, lngReturn) Then
'        Debug.Print "05 . 軟復 Soft Reconnect ( EQUIPMENT_UNLOCK ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            ' OK !
'            If UCase(strStatus) = "S" Then
'                gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
'                                        " SET CMOPENDATE=" & O2Date(RightNow) & _
'                                        ",UPDEN='CABLESOFT'" & _
'                                        ",UPDTIME='" & GetDTString(RightNow) & "'" & _
'                                        " WHERE FACISNO='" & Replace(str_CM_MAC, ":", "") & "'"
'            End If
''            With rsData
''                .Fields("CMOpenDate") = Format(RightNow, strYMDHMS)
''                .Fields("UpdEn") = "CABLESOFT"
''                .Fields("UpdTime") = GetDTString(RightNow)
''                .Update
''            End With
'        End If
'    Else
'        ShowMsg
'    End If
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_05"
'End Function
'
'' 乙、 06 cm 出庫 由設備資料選定後按照出庫所需參數取出SO004的相關資料與畫面設定狀態，指令執行結果(第二道查詢結果)
''   若成功需刪除so306該筆序號資料 (此功能使用時機為測試時使用，例行作業需作在cm 出庫作業時需傳送此指令)
'Public Function Cmd_06(str_CM_MAC As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean    ' 2. CM出庫 [06]
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123000200001','','','00:08:0E:AC:A8:9C',:vMsgID);
'
'    strCmdName = "06. CM 出庫"
'    strCompCode = GetSvcPvd
'    strCMmac = str_CM_MAC
'
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM,MODEMMAC" & _
'                            ",RETCODE,RETMSG,STATUSCODE,FAILURECODE,Q_STATUS_CODE,Q_FAILURE_CODE" & _
'                            ",CMDFLAG,SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "06", "EQUIPMENT_REMOVE"
'
'    With rs309
'            If .State > 0 Then .Update
'    End With
'
'    If Call_IEL_06(strCompCode, strCmdSno, "", "", strCMmac, lngMsgID, lngReturn) Then
'        Debug.Print "06. CM 出庫 Remove from inventory ( EQUIPMENT_REMOVE ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            ' OK !
'        End If
'    Else
'        ShowMsg
'    End If
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_06"
'End Function

' 丁、 07 速率升降級 由設備資料選定後按照服務設定所需參數 並取出 SO004 CM 速率資料顯示於畫面上為預設值，
'   下拉清單由 CM 代碼檔帶入傳送指令時以新選定的速率傳送指令，指令執行結果(第二道查詢結果) 若成功需要回填 SO004 的 CM 速率

' 戊、 07 變更 CPE MAC 由設備資料選定後按照服務設定所需參數 並取出該設備 CPE MAC 資料顯示於畫面 Gird 上，作新舊對照，
'   傳送指令時需將以有異動新的 MAC 異動一組需要傳送一組指令，異動 N 組區傳 N 次指令與沒異動的 MAC 一並組成參數所需格式傳送，
'   指令執行結果(第二道查詢結果)若成功需要回填 CPE MAC 對應的資料
'Public Function Cmd_07(str_CM_MAC As String, _
'                                        str_Cust_ID As String, _
'                                        str_CM_Baud_Rate As String, _
'                                        str_CPE_MAC As String, _
'                                        str_CPE_MAC_NEW As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean ' 4. 速率升降級 [07]
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123001100001','','','00:08:0E:AC:A8:9C','6666','284K','A0:00:00:00:00:01','FF:00:00:00:00:01',:vMsgID
''    Dim str_Cust_ID As String
''    Dim str_CM_Baud_Rate As String ' CM BaudRate : SO004.CMBaudRate
''    Dim str_CPE_MAC As String
''    Dim str_CPE_MAC_NEW As String
'
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM,MODEMMAC," & _
'                            "INTERNETTYPENAME,NUMDYNAIP,NUMSTATIP,CPEMAC,CPEMACNEW," & _
'                            "RETCODE,RETMSG,STATUSCODE,FAILURECODE,Q_STATUS_CODE,Q_FAILURE_CODE," & _
'                            "CMDFLAG,SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strCmdName = "07 . 速率升降級 / 變更CPE MAC"
''    strCmdSno = Get_Cmd_Sno
'    strCompCode = GetSvcPvd
'    strCMmac = str_CM_MAC ' strCMMAC = Mask(rsData("FaciSno") & "")
'
''    With frmSO1623A
''            str_Cust_ID = .ginCustId.Text
''            If .optCMChange(3).Value Then
''                str_CM_Baud_Rate = .gilBaudRate.GetDescription
''                str_CPE_MAC = ""
''                str_CPE_MAC_NEW = ""
''            Else
''                str_CPE_MAC = .cboCPE
''                str_CPE_MAC_NEW = .txtNewCPE
'''                str_CPE_MAC = Get_CPE_Mac(Val(str_Fix_IP_Count))
''            End If
''    End With
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "07", "SERVICE_UPDATE"
'
'    With rs309
'            If .State > 0 Then
'                .Fields("INTERNETTYPENAME") = str_CM_Baud_Rate
'                .Fields("CPEMAC") = str_CPE_MAC
'                .Fields("CPEMACNEW") = str_CPE_MAC_NEW
'                .Update
'            End If
'    End With
'
'    If Call_IEL_07(strCompCode, strCmdSno, "", "", strCMmac, str_Cust_ID, str_CM_Baud_Rate, _
'                            str_CPE_MAC, str_CPE_MAC_NEW, lngMsgID, lngReturn) Then
'        Debug.Print "07 . 速率升降級 / 變更CPE MAC ( Baud rate Upgrade / Downgrate / Change CPE MAC ( SERVICE_UPDATE ) ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            ' OK !
'            If UCase(strStatus) = "S" Then
'                If str_CPE_MAC_NEW = Empty Then
'                    gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
'                                                " SET CMBAUDRATE='" & str_CM_Baud_Rate & "'" & _
'                                                ",UPDEN='CABLESOFT'" & _
'                                                ",UPDTIME='" & GetDTString(RightNow) & "'" & _
'                                                " WHERE FACISNO='" & Replace(str_CM_MAC, ":", "") & "'"
'                Else
'                    gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
'                                                " SET CPEMAC='" & str_CPE_MAC_NEW & "'" & _
'                                                ",UPDEN='CABLESOFT'" & _
'                                                ",UPDTIME='" & GetDTString(RightNow) & "'" & _
'                                                " WHERE FACISNO='" & Replace(str_CM_MAC, ":", "") & "'"
'                End If
'            End If
''            If frmSO1623A.optCMChange(3).Value Then
''                With rsData
''                    .Fields("CMBaudRate") = str_CM_Baud_Rate
''                    .Fields("UpdEn") = "CABLESOFT"
''                    .Fields("UpdTime") = GetDTString(RightNow)
''                    .Update
''                End With
''            'ElseIf frmSO1623A.optCMChange(3).Value Then
''            Else
''                With rsData
''                    .Fields("CPEMAC") = str_CPE_MAC_NEW
''                    .Fields("UpdEn") = "CABLESOFT"
''                    .Fields("UpdTime") = GetDTString(RightNow)
''                    .Update
''                End With
''            End If
'        End If
'    Else
'        ShowMsg
'    End If
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_07"
'End Function

' 戊、 08 CP服務增加，由設備資料選定後需該設備參考號屬於 EMTA 時，取 SO004 相關資料且需用該序號關聯檢核CP服務，
'   是否有安裝中或正常的 CP 門號取出指令所需參數後傳送 (若對應多組門號取出一組參數即可)，
'   執行結果(第二道查詢結果)若成功需要回填SO004 該 CP 門號的開機時間

' 丙、 08 申請 IP : 由設備資料選定後按照服務設定所需參數 並取出 SO004 固定 IP 與動態IP帶入畫面, 有權限的人員可修改此兩個欄位,
' 但是不可小於設備之欄位帶出來的數目總數(固定IP與動態IP需區分開來,不可合計) ，
' 指令執行結果(第二道查詢結果) 若成功需要回填SO004對應的 IP 欄位值 '20050719 Lydia Modify

'Public Function Cmd_08(str_CM_MAC As String, _
'                                        str_Cust_ID As String, _
'                                        str_VOIP_Type As String, _
'                                        str_Media_Gateway As String, _
'                                        str_IP_Address As String, _
'                                        str_Dyna_IP_Count As String, _
'                                        str_Fix_IP_Count As String, _
'                                        str_CPE_MAC As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean  ' 5. CP 服務增加 [08]
'
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123000900001','','','00:08:0E:AC:A8:9C','6666','','','','','2','1',:vMsgID
'    ' Sample : '01','00120051123000900004','','','00:08:0E:15:C2:14','26664','VOIP','BSTG01','10.228.0.2','10.228.0.2','2','1',:vMsgID
'
''    Dim str_Cust_ID As String
''    Dim str_VOIP_Type As String ' VOIP Type : 有CP 服務時 需傳送 CD043.VoiceType ( N )
''    Dim str_Media_Gateway As String ' Media Gateway : 有CP 服務時 SO049.Gateway CP 服務由門號決定 ( Gateway ) ( N )
''    Dim str_IP_Address As String ' IP Address : 有CP 服務時  SO004.IPAddress 當品名參考號屬於EMTA 5 ( N )
''    Dim str_Dyna_IP_Count As String ' Number Dynamic IP : SO004.DynIPCount 需考慮增加或減少，以最後總數傳送，不應該 < 0 ( N )
''    Dim str_Fix_IP_Count As String ' Number Static IP : SO004.FixIPCount 需考慮增加或減少，以最後總數傳送，不應該 < 0 ( N )
''    Dim str_CPE_MAC As String ' CPE MAC Address : SO004C.CPEMAC ( N )
'
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM,MODEMMAC,VOIPTYPE," & _
'                            "MEDIAGATEWAY,IPADDRESS,NUMDYNAIP,NUMSTATIP,CPEMAC,RETCODE,RETMSG," & _
'                            "STATUSCODE,FAILURECODE,Q_STATUS_CODE,Q_FAILURE_CODE,CMDFLAG," & _
'                            "SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strCmdName = "08. CP裝機 (分機) / 申請IP"
'    'strCmdSno = Get_Cmd_Sno
'    strCompCode = GetSvcPvd
'    'strCMMAC = Mask(rsData("FaciSno") & "")
'    strCMmac = str_CM_MAC
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
''    With frmSO1623A
''            'str_Cust_ID = .ginCustId.Text
''            If .sstData.Tab = 1 Then
''                ' 申請IP Add Fix IP
''                str_VOIP_Type = ""
''                str_Media_Gateway = ""
''                str_IP_Address = ""
''                str_Dyna_IP_Count = Val(.txtDynaIP2)
''                str_Fix_IP_Count = Val(.txtFixIP2)
''                str_CPE_MAC = Get_CPE_Mac(Val(str_Fix_IP_Count))
''            Else
''                ' CP裝機(分機) ADD CP Installation
''                str_VOIP_Type = IIf(.IsCMCP, Get_VOIP_Type(rsData("ModelCode") & ""), "")
''                str_Media_Gateway = Get_Media_Gateway(rsData("FaciSno") & "") ' rsData("Gateway") & ""
''                str_IP_Address = rsData("IPAddress") & ""
''                str_Dyna_IP_Count = "" 'rsData("DynIPCount") & ""
''                str_Fix_IP_Count = "" ' rsData("FixIPCount") & ""
''                str_CPE_MAC = "" ' Get_CPE_Mac
''            End If
''    End With
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "08", "SERVICE_CREATE"
'
'    With rs309
'            If .State > 0 Then
'                .Fields("VOIPTYPE") = str_VOIP_Type
'                .Fields("MEDIAGATEWAY") = str_Media_Gateway
'                .Fields("IPADDRESS") = str_IP_Address
'                .Fields("NUMDYNAIP") = Val(str_Dyna_IP_Count)
'                .Fields("NUMSTATIP") = Val(str_Fix_IP_Count)
'                .Fields("CPEMAC") = str_CPE_MAC
'                .Update
'            End If
'    End With
'
'    If Call_IEL_08(strCompCode, strCmdSno, "", "", strCMmac, str_Cust_ID, str_VOIP_Type, str_Media_Gateway, _
'                            str_IP_Address, str_Dyna_IP_Count, str_Fix_IP_Count, str_CPE_MAC, lngMsgID, lngReturn) Then
'        Debug.Print "08 . CP裝機 (分機) / 申請IP ADD CP Installation / Add Fix IP ( SERVICE_CREATE ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            If UCase(strStatus) = "S" Then
'                If str_Dyna_IP_Count = Empty Then
'                    ' 若成功需要回填SO004 該cp門號的開機時間
'                    gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
'                                                " SET CMOPENDATE=" & O2Date(RightNow) & _
'                                                ",UPDEN='CABLESOFT'" & _
'                                                ",UPDTIME='" & GetDTString(RightNow) & "'" & _
'                                                " WHERE REFACISNO='" & Replace(str_CM_MAC, ":", "") & "'" & _
'                                                " AND COMPCODE=" & strCompCode & _
'                                                " AND CUSTID=" & str_Cust_ID & _
'                                                " AND SERVICETYPE='P'" & _
'                                                " AND PRDATE IS NULL"
'    '                                           " WHERE FACISNO='" & Replace(str_CM_MAC, ":", "") & "'"
'                Else
'                    ' 若成功需要回填SO004對應的 IP 欄位值
'                    gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
'                                                " SET DYNIPCOUNT=" & str_Dyna_IP_Count & _
'                                                ",FIXIPCOUNT=" & str_Fix_IP_Count & _
'                                                ",UPDEN='CABLESOFT'" & _
'                                                ",UPDTIME='" & GetDTString(RightNow) & "'" & _
'                                                " WHERE FACISNO='" & Replace(str_CM_MAC, ":", "") & "'"
'                End If
'            End If
''            If frmSO1623A.sstData.Tab = 0 Then
''                ' 若成功需要回填SO004 該cp門號的開機時間
''                gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
''                                            " SET CMOPENDATE=" & O2Date(RightNow) & _
''                                            ",UPDEN='CABLESOFT'" & _
''                                            ",UPDTIME='" & GetDTString(RightNow) & "'" & _
''                                            " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
''                                            " AND COMPCODE=" & rsData("CompCode") & _
''                                            " AND CUSTID=" & rsData("CustID") & _
''                                            " AND SERVICETYPE='P'" & _
''                                            " AND PRDATE IS NULL"
''            Else
''                ' 若成功需要回填SO004對應的 IP 欄位值
''                With rsData
''                    .Fields("DynIPCount") = Val(str_Dyna_IP_Count)
''                    .Fields("FixIPCount") = Val(str_Fix_IP_Count)
''                    .Fields("UpdEn") = "CABLESOFT"
''                    .Fields("UpdTime") = GetDTString(RightNow)
''                    .Update
''                End With
''            End If
'        End If
'    Else
'        ShowMsg
'    End If
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_08"
'End Function


' 己、 09 CP服務刪除，由設備資料選定後需該設備參考號屬於 EMTA 時，取 SO004 相關資料且需用該序號關聯檢核 CP 服務，
'   (門號狀態不判定) 的 CP 門號取出指令所需參數後傳送(若對應多組門號取出一組參數即可，
'   執行結果(第二道查詢結果)若成功需要回填SO004 該CP門號的關機時間

' 己、 09 取消IP: 由設備資料選定後按照服務設定所需參數 並取出SO004固定IP 與動態IP帶入畫面,
'   有權限的人員可修改此兩個欄位,但是不可大於設備之欄位帶出來的數目總數(固定IP與動態IP需區分開來,不可合計) ，
'   指令執行結果(第二道查詢結果)若成功需要回填SO004對應的ip 欄位值
'Public Function Cmd_09(str_CM_MAC As String, _
'                                        str_Cust_ID As String, _
'                                        str_VOIP_Type As String, _
'                                        str_Dyna_IP_Count As String, _
'                                        str_Fix_IP_Count As String, _
'                                        str_CPE_MAC As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean ' 6. CP 服務刪除 [09]
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123001000001','','','00:08:0E:AC:A8:9C','6666','','1','1','A0:00:00:00:00:01',:vMsgID
'    ' Sample : '01','00120051123001000004','','','00:08:0E:15:C2:14','26664','','1','1','',:vMsgID
'
''    Dim str_Cust_ID As String
''    Dim str_VOIP_Type As String ' VOIP Type : 有CP 服務時 需傳送 CD043.VoiceType ( N )
''    Dim str_Dyna_IP_Count As String ' Number Dynamic IP : SO004.DynIPCount 需考慮增加或減少，以最後總數傳送，不應該 < 0 ( N )
''    Dim str_Fix_IP_Count As String ' Number Static IP : SO004.FixIPCount 需考慮增加或減少，以最後總數傳送，不應該 < 0 ( N )
''    Dim str_CPE_MAC As String ' CPE MAC Address : SO004C.CPEMAC ( N )
'
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM,MODEMMAC,VOIPTYPE," & _
'                            "NUMDYNAIP,NUMSTATIP,CPEMAC,RETCODE,RETMSG,STATUSCODE,FAILURECODE," & _
'                            "Q_STATUS_CODE,Q_FAILURE_CODE,CMDFLAG,SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strCmdName = "09 . 取消IP / CP Only 拆機 (分機)"
'    'strCmdSno = Get_Cmd_Sno
'    strCompCode = GetSvcPvd
'    strCMmac = str_CM_MAC ' strCMMAC = Mask(rsData("FaciSno") & "")
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
''    With frmSO1623A
''            str_Cust_ID = .ginCustId.Text
''            If .optCMChange(5).Value Then
''                ' 取消IP Minus Fix IP
''                str_VOIP_Type = "" ' * : IIf(.IsCMCP, Get_VOIP_Type(rsData("ModelCode") & ""), "")
''                str_Dyna_IP_Count = Val(.txtDynaIP3)
''                str_Fix_IP_Count = Val(.txtFixIP3)
''                str_CPE_MAC = Get_CPE_Mac(Val(str_Fix_IP_Count))
''            Else
''                ' CP Only 拆機(分機) Remove CP Service
''                str_VOIP_Type = "" ' * : IIf(.IsCMCP, Get_VOIP_Type(rsData("ModelCode") & ""), "")
''                str_Dyna_IP_Count = "" ' Optional : rsData!DynIPCount & ""
''                str_Fix_IP_Count = "" ' Optional : rsData!FixIPCount & ""
''                str_CPE_MAC = "" ' Optional : Get_CPE_Mac
''            End If
''    End With
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "09", "SERVICE_REMOVE"
'
'    With rs309
'            If .State > 0 Then
'                .Fields("VOIPTYPE") = str_VOIP_Type
'                .Fields("NUMDYNAIP") = Val(str_Dyna_IP_Count)
'                .Fields("NUMSTATIP") = Val(str_Fix_IP_Count)
'                .Fields("CPEMAC") = str_CPE_MAC
'                .Update
'            End If
'    End With
'
'    If Call_IEL_09(strCompCode, strCmdSno, "", "", strCMmac, str_Cust_ID, str_VOIP_Type, _
'                            str_Dyna_IP_Count, str_Fix_IP_Count, str_CPE_MAC, lngMsgID, lngReturn) Then
'        Debug.Print "09 . 取消IP / CP Only 拆機 (分機) Minus Fix IP / Add Fix IP / Remove CP Service ( SERVICE_REMOVE ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            If UCase(strStatus) = "S" Then
'                If str_Dyna_IP_Count = Empty Then
'                    ' 若成功需要回填SO004 該CP門號的關機時間
'                    gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
'                                                " SET CMCLOSEDATE=" & O2Date(RightNow) & _
'                                                ",UPDEN='CABLESOFT'" & _
'                                                ",UPDTIME='" & GetDTString(RightNow) & "'" & _
'                                                " WHERE REFACISNO='" & Replace(str_CM_MAC, ":", "") & "'" & _
'                                                " AND COMPCODE=" & strCompCode & _
'                                                " AND CUSTID=" & str_Cust_ID & _
'                                                " AND SERVICETYPE='P'" & _
'                                                " AND PRDATE IS NULL"
'                Else
'                    ' 若成功需要回填 SO004對應的 IP 欄位值
'                    gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
'                                                " SET DYNIPCOUNT=" & str_Dyna_IP_Count & _
'                                                ",FIXIPCOUNT=" & str_Fix_IP_Count & _
'                                                ",UPDEN='CABLESOFT'" & _
'                                                ",UPDTIME='" & GetDTString(RightNow) & "'" & _
'                                                " WHERE FACISNO='" & Replace(str_CM_MAC, ":", "") & "'"
'                End If
'            End If
''            If frmSO1623A.sstData.Tab = 0 Then
''                ' 若成功需要回填SO004 該CP門號的關機時間
''                gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
''                                            " SET CMCLOSEDATE=" & O2Date(RightNow) & _
''                                            ",UPDEN='CABLESOFT'" & _
''                                            ",UPDTIME='" & GetDTString(RightNow) & "'" & _
''                                            " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
''                                            " AND COMPCODE=" & rsData("CompCode") & _
''                                            " AND CUSTID=" & rsData("CustID") & _
''                                            " AND SERVICETYPE='P'" & _
''                                            " AND PRDATE IS NULL"
''            Else
''                If frmSO1623A.optCMChange(5).Value Then
''                    ' 若成功需要回填 SO004對應的 IP 欄位值
''                    With rsData
''                        .Fields("DynIPCount") = Val(str_Dyna_IP_Count)
''                        .Fields("FixIPCount") = Val(str_Fix_IP_Count)
''                        .Fields("UpdEn") = "CABLESOFT"
''                        .Fields("UpdTime") = GetDTString(RightNow)
''                        .Update
''                    End With
''                End If
''            End If
'        End If
'    Else
'        ShowMsg
'    End If
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_09"
'End Function

' 庚、 10 網路編號 Node 與 IP 異動 以該裝機地址取出 Node 或 網編 (系統參數決定)若屬於 EMTA 設備時需要加取IP資料出來，
'   若該客戶(不分服務別)有移機工單未完工的則移機中取用選項可提供選擇，
'   選擇後檢查網編或node與舊值是否相同，若不同以新 Node (網編) 取一組ip值出來作新值後傳送，指令執行結果(第二道查詢結果)
'   若成功需要回填SO004對應的 IPAddress 欄位值
'   (Node代碼檔,SO016.CircuitNo,IP為004 , SO048,UseFlag, CircuitNo,左邊為空或為原值,則右邊不能選 )

'   移機中定義 : SO002.WipCode3=11 & SO007.PRCODE Ref-> CD007.RefNo=3 , ServiceType = ?

'Public Function Cmd_10(str_CM_MAC As String, _
'                                        str_Cust_ID As String, _
'                                        str_HFC_Node As String, _
'                                        str_IP_Address As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean ' 8. 網路編號 / NODE or Voice IP [10]
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123001200001','','','00:08:0E:AC:A8:9C','','Z28',:vMsgID
'    ' EMTA IP PHONE Priv IP 設定 (移機) / CM IP  CPE IP 設定 (移機)
'    ' 10 . EMTA IP PHONE Priv IP 設定 (移機)
'    ' 10 . CM IP  CPE IP 設定 (移機)
'
''    Dim str_HFC_Node As String
''    Dim str_IP_Address As String
'
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM," & _
'                            "MODEMMAC,HFCNODE,IPADDRESS,RETCODE,RETMSG," & _
'                            "STATUSCODE,FAILURECODE,Q_STATUS_CODE,Q_FAILURE_CODE," & _
'                            "CMDFLAG,SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strCmdName = "10. 網路編號 / NODE / VOIP"
''    strCmdSno = Get_Cmd_Sno
'    strCompCode = GetSvcPvd
'    strCMmac = str_CM_MAC ' strCMMAC = Mask(rsData("FaciSno") & "")
'
''    With frmSO1623A
''            str_HFC_Node = .cboNode
''            str_IP_Address = .txtNewIP
''    End With
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "10", "HFC_NODE_UPDATE"
'
'    With rs309
'            If .State > 0 Then
'                .Fields("HFCNODE") = str_HFC_Node
'                .Fields("IPADDRESS") = str_IP_Address
'                .Update
'            End If
'    End With
'
'    If Call_IEL_10(strCompCode, strCmdSno, "", "", strCMmac, str_IP_Address, str_HFC_Node, lngMsgID, lngReturn) Then
'        Debug.Print "10. 網路編號 / NODE / VOIP ( HFC_NODE_UPDATE ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            ' 若成功需要回填SO004對應的 IPAddress 欄位值
'            If UCase(strStatus) = "S" Then
'                If str_IP_Address <> Empty Then
'                    frmSO1623A.blnGetIP = False
'                    gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
'                                                " SET IPADDRESS='" & str_IP_Address & "'" & _
'                                                ",UPDEN='CABLESOFT'" & _
'                                                ",UPDTIME='" & GetDTString(RightNow) & "'" & _
'                                                " WHERE REFACISNO='" & Replace(str_CM_MAC, ":", "") & "'" & _
'                                                " AND COMPCODE=" & strCompCode & _
'                                                " AND CUSTID=" & str_Cust_ID & _
'                                                " AND SERVICETYPE='I'" & _
'                                                " AND PRDATE IS NULL"
'                End If
'            End If
'        End If
'    Else
'        ShowMsg
'    End If
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_10"
'End Function

'Private Function GetIPAddress(strReFaciSno As String) As String
'  On Error Resume Next
'    With frmSO1623A
'        GetIPAddress = RPxx(gcnGi.Execute("SELECT IPADDRESS FROM " & GetOwner & "SO004" & _
'                                                                " WHERE FACISNO='" & strReFaciSno & "'" & _
'                                                                " AND COMPCODE=" & .gilCompCode.GetCodeNo & " " & _
'                                                                " AND CUSTID=" & .ginCustId.Text & _
'                                                                " AND SERVICETYPE='I'").GetString(adClipString, , "", , ""))
'        If Err.Number <> 0 Then GetIPAddress = ""
'    End With
'End Function

'txtIPAddress.Text = GetIPAddress(lngAddrNo, strCircuitNO)
'.uOldNo = GetRsValue("Select NodeNo From " & GetOwner & "SO014 Where AddrNo = " & lngAddrNo & " And CompCode = " & gCompCode) & ""
'.uOldNo = GetRsValue("Select CircuitNO From " & GetOwner & "SO014 Where AddrNo = " & lngAddrNo & " And CompCode = " & gCompCode) & ""
' "Select IPAddress,UseFlag,CircuitNo From " & GetOwner & "SO048 Where " & strWhere & " CompCode=" & gCompCode & " and UseFlag=0 Order By IPAddress"

'Private Function GetIPAddress(ByVal lngAddrNo As Long, _
'    ByRef strCircuitNO As String) As String
'    On Error Resume Next
'    Dim strField As String
'        If intEMTAIPType = 2 Then
'            strField = "NodeNo"
'        Else
'            strField = "CircuitNo"
'        End If
'        strCircuitNO = GetRsValue("Select " & strField & " From " & GetOwner & "SO014 Where AddrNo = " & lngAddrNo & _
'                        " And CompCode = " & gCompCode) & ""
'
'        GetIPAddress = GetRsValue("Select IPAddress from " & GetOwner & "SO048 where " & _
'                        "CircuitNo = '" & strCircuitNO & "' and CompCode=" & gCompCode & " and UseFlag=0")
'End Function
'
'Private Function GetIPAddress() As String
'    On Error Resume Next
'    Dim strField As String
'        If intFaciRefNo = 5 Then
'            If intEMTAIPType = 2 Then
'                strField = "NodeNo"
'            Else
'                strField = "CircuitNo"
'            End If
'            GetIPAddress = GetRsValue("Select IPAddress from " & GetOwner & "SO048 where " & _
'                            "CircuitNo = (Select " & strField & " From " & GetOwner & "SO014 Where AddrNo = " & lngAddrNo & _
'                            " And CompCode = " & gCompCode & ") and CompCode=" & gCompCode & " and UseFlag=0") & ""
'        End If
'End Function

'        If intFaciRefNo = 5 And txtIPAddress <> "" Then
'            '如果為EMTA,需將UseFlag 設成1 93/03/23
'            gcnGi.Execute "Update " & GetOwner & "SO048 Set UseFlag=1 Where IPAddress='" & txtIPAddress & "' And CompCode =" & gCompCode
'        ElseIf intFaciRefNo = 6 And txtFaciSNo <> "" Then
'            '如果為CP電話號碼,需將UseFlag 設成1 93/03/16
'            gcnGi.Execute "Update " & GetOwner & "SO049 Set UseFlag=1 Where CPNumber='" & txtFaciSNo & "' And CompCode =" & gCompCode
'        End If

'Public Function Cmd_11(ByRef rsData As ADODB.Recordset) As Boolean ' 拆舊/換新(只限同TYPE更換)
'  On Error GoTo ChkErr
    ' 11 . 拆舊 / 換新 (只限同TYPE更換)
    ' Replace damaged CM ( EQUIPMENT_SWAP )
    ' Para :
    '           SERVICE_PROVIDER
    '           MESSAGE_LOG_IDENTIFIER1
    '           MODEM_MAC_ADDRESS
    '           MTA_MAC_ADDRESS ( Optional )
    '           EQUIPMENT_TYPE
    '           EQUIPMENT_STATUS
    '           MODEM_MAC_ADDRESS_NEW
    '           MTA_MAC_ADDRESS_NEW  ( Optional )
    '           ACCOUNT_NUMBER
    '           EQUIPMENT_TYPE
    ' Sample : 01','00120051123000600001','','','00:08:0E:AC:A8:9C','00:0E:5C:60:61:7E','','','INACTIVE','MotoSB510X',:vMsgID
'    CREATE OR REPLACE FUNCTION
'        SF_Call_IEL_11(
'            p_Service_Provider       IN     VARCHAR2,
'            p_Msg_ID_1       IN     VARCHAR2,
'            p_Msg_ID_2       IN     VARCHAR2,
'            p_Msg_ID_3       IN     VARCHAR2,
'            p_Modem_MAC      IN     VARCHAR2,
'            p_Modem_MAC_New  IN     VARCHAR2,
'            p_MTA_MAC        IN     VARCHAR2,
'            p_MTA_MAC_New        IN     VARCHAR2,
'            p_CustID             IN     VARCHAR2,
'            p_Equip_Type         IN     VARCHAR2,
'            p_Msg_ID         OUT    PLS_INTEGER
'            ) RETURN PLS_INTEGER
'        Call_IEL_11 "ServiceProvider", "MsgID1", "", "", "ModemMac", "ModemMacNew", _
'                                "MTAMac", "MTAMacNew ", "CustID", "EquipType", "MsgID", "Return"
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_11"
'End Function
'
'Public Function Cmd_12(str_CM_MAC As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean  ' 7. CM 重設 [12]
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123000500001','','','00:08:0E:AC:A8:9C',:vMsgID
'
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM," & _
'                            "MODEMMAC,RETCODE,RETMSG,STATUSCODE,FAILURECODE," & _
'                            "Q_STATUS_CODE,Q_FAILURE_CODE,CMDFLAG," & _
'                            "SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strCmdName = "12 . 重設 CM Reset"
''    strCmdSno = Get_Cmd_Sno
'    strCompCode = GetSvcPvd
'    strCMmac = str_CM_MAC ' strCMMAC = Mask(rsData("FaciSno") & "")
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "12", "EQUIPMENT_RESET"
'
'    With rs309
'            If .State > 0 Then .Update
'    End With
'
'    If Call_IEL_12(strCompCode, strCmdSno, "", "", strCMmac, lngMsgID, lngReturn) Then
'        Debug.Print "12 . 重設 Reset Modem Parameter ( EQUIPMENT_RESET ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            ' OK !
'            If UCase(strStatus) = "S" Then
'                gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
'                                        " SET CMOPENDATE=" & O2Date(RightNow) & _
'                                        ",UPDEN='CABLESOFT'" & _
'                                        ",UPDTIME='" & GetDTString(RightNow) & "'" & _
'                                        " WHERE FACISNO='" & Replace(str_CM_MAC, ":", "") & "'"
'            End If
''            With rsData
''                .Fields("CMOpenDate") = Format(RightNow, strYMDHMS)
''                .Fields("UpdEn") = "CABLESOFT"
''                .Fields("UpdTime") = GetDTString(RightNow)
''                .Update
''            End With
'        End If
'    Else
'        ShowMsg
'    End If
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_12"
'End Function
'
'' 甲、 13 cm KINDCLASS 查詢 由設備資料選定後按照服務設定所需參數，指令執行結果(第二道查詢結果)
''   若成功，需要傳送第三道指令取回訊息，彈出視窗顯示並紀錄自log內。
'
'Public Function Cmd_13(str_CM_MAC As String, _
'                                        str_Cust_ID As String, _
'                                        str_MTA_MAC As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean ' 1. 查詢 CM 狀況 [13]
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123001300001','','','00:08:0E:AC:A8:9C','','6666',:vMsgID
'
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM," & _
'                            "MODEMMAC,MTAMAC,CLIENTCLASSNAME,RETCODE,RETMSG," & _
'                            "STATUSCODE,FAILURECODE,Q_STATUS_CODE,Q_FAILURE_CODE," & _
'                            "CMDFLAG,SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strCmdName = "13 . 查詢CM狀況"
''    strCmdSno = Get_Cmd_Sno
'    strCompCode = GetSvcPvd
'    strCMmac = str_CM_MAC ' strCMMAC = Mask(rsData("FaciSno") & "")
'    strMTAMAC = str_MTA_MAC ' strMTAMAC = Get_EMTA_Mac(rsData("FaciSNo") & "")
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "13", "GET_CONFIGURATION"
'
'    With rs309
'            If .State > 0 Then
'                .Fields("MTAMAC") = strMTAMAC
'                .Update
'            End If
'    End With
'
'    blnNoShwMsg = True
'
'    If Call_IEL_13(strCompCode, strCmdSno, "", "", strCMmac, strMTAMAC, str_Cust_ID, lngMsgID, lngReturn) Then
'        Debug.Print "13 . 查詢CM狀況 CM Status Inquiry ( GET_CONFIGURATION ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            If UCase(strStatus) = "S" Then
'                If Pooling(frmSO1623A.pbr1, "Cmd_14") Then
'                    ' OK !
'                Else
'                    ' No OK !
'                End If
'            Else
'                If Not blnQuery Then ShowMsg
'            End If
'        End If
'    Else
'        ShowMsg
'    End If
'
'    blnNoShwMsg = False
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_13"
'    blnNoShwMsg = False
'End Function
'
'Public Function Cmd_14(blnShowMsg As Boolean) As Boolean  ' 14 . 查詢CM狀況結果 CM Status Inquiry Result ( GET_CONFIGURATION_RESULT )
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123001300001',:vClientClass,:vMsgID
'    Dim strClientClass As String
'    If Call_IEL_14(strCompCode, strCmdSno, strClientClass, lngMsgID, lngReturn) Then
'        Debug.Print "14 . 查詢CM狀況結果 CM Status Inquiry Result ( GET_CONFIGURATION_RESULT ) 送出完成! MsgID : " & lngMsgID
'        Cmd_14 = True
'        With rs309
'                If .State > 0 Then .Fields("CLIENTCLASSNAME") = strClientClass
'        End With
'        ShowMsg
'    Else
'        Cmd_14 = False
'        Debug.Print "14 . 查詢CM狀況結果 CM Status Inquiry Result ( GET_CONFIGURATION_RESULT ) 送出失敗! MsgID : " & lngMsgID
'        If blnShowMsg Then ShowMsg
'    End If
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_14"
'End Function
'
''乙、15 CM 狀態查詢由設備資料選定後按照服務設定所需參數，指令執行結果(第二道查詢結果)
''   若成功，需要傳送第三道指令取回訊息，彈出視窗顯示並紀錄自log內。
'
'Public Function Cmd_15(str_CM_MAC As String, _
'                                        str_Sno As String, _
'                                        Optional str_Cmd_Sno As String = "", _
'                                        Optional blnQuery As Boolean = False) As Boolean ' 2. 查詢 CM 參數 [15]
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123001400001','','','00:08:0E:AC:A8:9C',:vMsgID
'
'    str309SQL = "SERVICEPROVIDER,MSGCODE,MSGNAME,CMDSEQNO,ACCOUNTNUM," & _
'                            "MODEMMAC,RETCODE,RETMSG,Q_IPADDRESS,Q_MACADDRESS," & _
'                            "Q_LEASETIME,Q_INTERFACE,Q_ADDRESSSTATE," & _
'                            "Q_IPLEASESTARTTIME,Q_IPLEASEINITSTARTTIME," & _
'                            "Q_IPLEASEPROTOCOL,Q_IPLEASEREMOTEID," & _
'                            "Q_SUBNETMASK,Q_DOMAINSERVER,Q_HOSTNAME," & _
'                            "Q_TFTPSERVERNAME,Q_BOOTFILENAME,Q_GATEWAYS," & _
'                            "STATUSCODE,FAILURECODE,Q_STATUS_CODE,Q_FAILURE_CODE," & _
'                            "CMDFLAG,SNO,EXECENTRY,EXECDATE,HOSTNAME,SOURCE"
'
'    strCmdName = "15 . 查詢CM參數"
''    strCmdSno = Get_Cmd_Sno
'    strCompCode = GetSvcPvd
'    strCMmac = str_CM_MAC ' strCMMAC = Mask(rsData("FaciSno") & "")
'
'    strStatus = ""
'    lngFailureCode = 0
'    lngReturn = 0
'
'    If blnQuery Then
'        strCmdSno = str_Cmd_Sno
'        GetRS309 str309SQL, strCmdSno
'        GoTo Qry
'    Else
'        strCmdSno = Get_Cmd_Sno
'    End If
'
'    Chk_Resend str_Cmd_Sno
'
'    Insert309 str_Sno, "15", "CM_INQUIRY"
'
'    With rs309
'            If .State > 0 Then .Update
'    End With
'
'    blnNoShwMsg = True
'
'    If Call_IEL_15(strCompCode, strCmdSno, "", "", strCMmac, lngMsgID, lngReturn) Then
'        Debug.Print "15 . 查詢CM參數 Message CM Inquiry ( CM_INQUIRY ) 送出完成! MsgID : " & lngMsgID
'Qry:
'        If Pooling(frmSO1623A.pbr1, "Qry_Result") Then
'            If UCase(strStatus) = "S" Then
'                If Pooling(frmSO1623A.pbr1, "Cmd_16") Then
'                    ' OK !
'                Else
'                    ' No OK !
'                End If
'            Else
'                If Not blnQuery Then ShowMsg
'            End If
'        End If
'    Else
'        ShowMsg
'    End If
'
'    blnNoShwMsg = False
'
'    With rs309
'            If .State > 0 Then
'                 Update309
'                .Update
'            End If
'    End With
'
'    ShowResult
'
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_15"
'    blnNoShwMsg = False
'End Function
'
'Public Function Cmd_16(blnShowMsg As Boolean) As Boolean ' 16 . 查詢CM參數結果 Message CM Inquiry Result ( CM_INQUIRY_RESULT )
'  On Error GoTo ChkErr
'    ' Sample : '01','00120051123001400001',:vIP_Address,:vMAC_Address,:vLease_Time,:vInterface,:vAddress_State,:vIP_Lease_Start_Time,:vIP_Lease_Init_Start_Time,:vIP_Lease_Protocol,:vIP_Lease_Remote_ID,:vSubnet_Mask,:vDomain_Server,:vHost_Name,:vT_FTP_Server_Name,    :vBoot_File_Name,:vGateways,:vStatic_Address_Name,:vStatic_Address_Desc,:vHardware_Address,:vHardware_Address_Type,:vIdentify_Client,:v_Static_IP,:v_Template,:vMsg_ID
'
'    Dim str_IP_Address As String
'    Dim str_MAC_Address As String
'    Dim str_Lease_Time As String
'    Dim str_Interface As String
'    Dim str_Address_State As String
'    Dim str_IP_Lease_Start_Time As String
'    Dim str_IP_Lease_Init_Start_Time As String
'    Dim str_IP_Lease_Protocol As String
'    Dim str_IP_Lease_Remote_ID As String
'    Dim str_Subnet_Mask As String
'    Dim str_Domain_Server As String
'    Dim str_Host_Name As String
'    Dim str_T_FTP_Server_Name As String
'    Dim str_Boot_File_Name As String
'    Dim str_Gateways As String
'    Dim str_Static_Address_Name As String
'    Dim str_Static_Address_Desc As String
'    Dim str_Hardware_Address As String
'    Dim str_Hardware_Address_Type As String
'    Dim str_Identify_Client As String
'    Dim str_Static_IP As String
'    Dim str_Template As String
'
'    If Call_IEL_16(strCompCode, strCmdSno, str_IP_Address, str_MAC_Address, str_Lease_Time, str_Interface, _
'                            str_Address_State, str_IP_Lease_Start_Time, str_IP_Lease_Init_Start_Time, _
'                            str_IP_Lease_Protocol, str_IP_Lease_Remote_ID, str_Subnet_Mask, _
'                            str_Domain_Server, str_Host_Name, str_T_FTP_Server_Name, str_Boot_File_Name, _
'                            str_Gateways, str_Static_Address_Name, str_Static_Address_Desc, _
'                            str_Hardware_Address, str_Hardware_Address_Type, str_Identify_Client, _
'                            str_Static_IP, str_Template, lngMsgID, lngReturn) Then
'
'        Debug.Print "16 . 查詢CM參數結果 Message CM Inquiry Result ( CM_INQUIRY_RESULT ) 送出完成! MsgID : " & lngMsgID
'        Cmd_16 = True
'        With rs309
'                If .State > 0 Then
'                    .Fields("Q_IPAddress") = str_IP_Address
'                    .Fields("Q_MacAddress") = str_MAC_Address
'                    .Fields("Q_LeaseTime") = str_Lease_Time
'                    .Fields("Q_Interface") = str_Interface
'                    .Fields("Q_AddressState") = str_Address_State
'                    .Fields("Q_IPLeaseStartTime") = str_IP_Lease_Start_Time
'                    .Fields("Q_IPLeaseInitStartTime") = str_IP_Lease_Init_Start_Time
'                    .Fields("Q_IPLeaseProtocol") = str_IP_Lease_Protocol
'                    .Fields("Q_IPLeaseRemoteID") = str_IP_Lease_Remote_ID
'                    .Fields("Q_SubnetMask") = str_Subnet_Mask
'                    .Fields("Q_DomainServer") = str_Domain_Server
'                    .Fields("Q_HostName") = str_Host_Name
'                    .Fields("Q_TFTPServerName") = str_T_FTP_Server_Name
'                    .Fields("Q_BootFileName") = str_Boot_File_Name
'                    .Fields("Q_Gateways") = str_Gateways
'                    .Fields("str_Static_Address_Name") = str_Static_Address_Name
'                    .Fields("str_Static_Address_Desc") = str_Static_Address_Desc
'                    .Fields("str_Hardware_Address") = str_Hardware_Address
'                    .Fields("str_Hardware_Address_Type") = str_Hardware_Address_Type
'                    .Fields("str_Identify_Client") = str_Identify_Client
'                    .Fields("str_Static_IP") = str_Static_IP
'                    .Fields("str_Template") = str_Template
'                End If
'        End With
'        ShowMsg
'    Else
'        Cmd_16 = False
'        Debug.Print "16 . 查詢CM參數結果 Message CM Inquiry Result ( CM_INQUIRY_RESULT ) 送出失敗! MsgID : " & lngMsgID
'        If blnShowMsg Then ShowMsg
'    End If
'
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Cmd_16"
'End Function

'Option Explicit
'
'Public Function Call_IEL_00(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByRef str_Status As String, _
'                                                    Optional ByRef lng_Failure_Code As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_00 As New ADODB.Command
'
'    With cmd_Call_IEL_00
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Status", adVarChar, adParamOutput, 4000, str_Status)
'        .Parameters.Append .CreateParameter("p_Failure_Code", adVarNumeric, adParamOutput, 2000, lng_Failure_Code)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_00"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         str_Status = .Parameters("p_Status").Value & ""
'         lng_Failure_Code = Val(.Parameters("p_Failure_Code").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_00 = (lng_Return = 0)
'
''    Call_IEL_00 = False
''    lng_Return = 1
''    lng_Failure_Code = 1000
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_00"
'    Call_IEL_00 = False
'End Function
'
'Public Function Call_IEL_01(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByVal str_MTA_MAC As String, _
'                                                    Optional ByVal str_Equip_Status As String, _
'                                                    Optional ByVal str_Equip_Type As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_01 As New ADODB.Command
'
'    With cmd_Call_IEL_01
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_MTA_MAC", adVarChar, adParamInput, 4000, str_MTA_MAC)
'        .Parameters.Append .CreateParameter("p_Equip_Status", adVarChar, adParamInput, 4000, str_Equip_Status)
'        .Parameters.Append .CreateParameter("p_Equip_Type", adVarChar, adParamInput, 4000, str_Equip_Type)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_01"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'
'         Call_IEL_01 = (lng_Return = 0)
'
'    End With
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_01"
'    Call_IEL_01 = False
'End Function
'
'Public Function Call_IEL_02(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByVal str_CustID As String, _
'                                                    Optional ByVal str_HFC_Node As String, _
'                                                    Optional ByVal str_CM_BaudRate As String, _
'                                                    Optional ByVal str_VOIP_Type As String, _
'                                                    Optional ByVal str_Media_Gateway As String, _
'                                                    Optional ByVal str_IP_Address As String, _
'                                                    Optional ByVal str_Dyna_IP_Num As String, _
'                                                    Optional ByVal str_Static_IP_Num As String, _
'                                                    Optional ByVal str_CPE_MAC As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_02 As New ADODB.Command
'
'    With cmd_Call_IEL_02
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_CustID", adVarChar, adParamInput, 4000, str_CustID)
'        .Parameters.Append .CreateParameter("p_HFC_Node", adVarChar, adParamInput, 4000, str_HFC_Node)
'        .Parameters.Append .CreateParameter("p_CM_BaudRate", adVarChar, adParamInput, 4000, str_CM_BaudRate)
'        .Parameters.Append .CreateParameter("p_VOIP_Type", adVarChar, adParamInput, 4000, str_VOIP_Type)
'        .Parameters.Append .CreateParameter("p_Media_Gateway", adVarChar, adParamInput, 4000, str_Media_Gateway)
'        .Parameters.Append .CreateParameter("p_IP_Address", adVarChar, adParamInput, 4000, str_IP_Address)
'        .Parameters.Append .CreateParameter("p_Dyna_IP_Num", adVarChar, adParamInput, 4000, str_Dyna_IP_Num)
'        .Parameters.Append .CreateParameter("p_Static_IP_Num", adVarChar, adParamInput, 4000, str_Static_IP_Num)
'        .Parameters.Append .CreateParameter("p_CPE_MAC", adVarChar, adParamInput, 4000, str_CPE_MAC)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_02"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_02 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_02"
'    Call_IEL_02 = False
'End Function
'
'Public Function Call_IEL_03(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByVal str_CustID As String, _
'                                                    Optional ByVal str_Equip_Status As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_03 As New ADODB.Command
'
'    With cmd_Call_IEL_03
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_CustID", adVarChar, adParamInput, 4000, str_CustID)
'        .Parameters.Append .CreateParameter("p_Equip_Status", adVarChar, adParamInput, 4000, str_Equip_Status)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_03"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_03 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_03"
'    Call_IEL_03 = False
'End Function
'
'Public Function Call_IEL_04(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByVal str_CustID As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_04 As New ADODB.Command
'
'    With cmd_Call_IEL_04
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_CustID", adVarChar, adParamInput, 4000, str_CustID)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_04"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_04 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_04"
'    Call_IEL_04 = False
'End Function
'
'Public Function Call_IEL_05(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByVal str_CustID As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_05 As New ADODB.Command
'
'    With cmd_Call_IEL_05
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_CustID", adVarChar, adParamInput, 4000, str_CustID)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_05"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_05 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_05"
'    Call_IEL_05 = False
'End Function
'
'Public Function Call_IEL_06(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_06 As New ADODB.Command
'
'    With cmd_Call_IEL_06
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_06"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_06 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_06"
'    Call_IEL_06 = False
'End Function
'
'Public Function Call_IEL_07(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByVal str_CustID As String, _
'                                                    Optional ByVal str_CM_BaudRate As String, _
'                                                    Optional ByVal str_CPE_MAC As String, _
'                                                    Optional ByVal str_CPE_MAC_NEW As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_07 As New ADODB.Command
'
'    With cmd_Call_IEL_07
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_CustID", adVarChar, adParamInput, 4000, str_CustID)
'        .Parameters.Append .CreateParameter("p_CM_BaudRate", adVarChar, adParamInput, 4000, str_CM_BaudRate)
'        .Parameters.Append .CreateParameter("p_CPE_MAC", adVarChar, adParamInput, 4000, str_CPE_MAC)
'        .Parameters.Append .CreateParameter("p_CPE_MAC_NEW", adVarChar, adParamInput, 4000, str_CPE_MAC_NEW)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_07"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_07 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_07"
'    Call_IEL_07 = False
'End Function
'
'Public Function Call_IEL_08(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByVal str_CustID As String, _
'                                                    Optional ByVal str_VOIP_Type As String, _
'                                                    Optional ByVal str_Media_Gateway As String, _
'                                                    Optional ByVal str_IP_Address As String, _
'                                                    Optional ByVal str_Dyna_IP_Num As String, _
'                                                    Optional ByVal str_Static_IP_Num As String, _
'                                                    Optional ByVal str_CPE_MAC As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_08 As New ADODB.Command
'
'    With cmd_Call_IEL_08
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_CustID", adVarChar, adParamInput, 4000, str_CustID)
'        .Parameters.Append .CreateParameter("p_VOIP_Type", adVarChar, adParamInput, 4000, str_VOIP_Type)
'        .Parameters.Append .CreateParameter("p_Media_Gateway", adVarChar, adParamInput, 4000, str_Media_Gateway)
'        .Parameters.Append .CreateParameter("p_IP_Address", adVarChar, adParamInput, 4000, str_IP_Address)
'        .Parameters.Append .CreateParameter("p_Dyna_IP_Num", adVarChar, adParamInput, 4000, str_Dyna_IP_Num)
'        .Parameters.Append .CreateParameter("p_Static_IP_Num", adVarChar, adParamInput, 4000, str_Static_IP_Num)
'        .Parameters.Append .CreateParameter("p_CPE_MAC", adVarChar, adParamInput, 4000, str_CPE_MAC)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_08"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_08 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_08"
'    Call_IEL_08 = False
'End Function
'
'Public Function Call_IEL_09(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByVal str_CustID As String, _
'                                                    Optional ByVal str_VOIP_Type As String, _
'                                                    Optional ByVal str_Dyna_IP_Num As String, _
'                                                    Optional ByVal str_Static_IP_Num As String, _
'                                                    Optional ByVal str_CPE_MAC As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_09 As New ADODB.Command
'
'    With cmd_Call_IEL_09
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_CustID", adVarChar, adParamInput, 4000, str_CustID)
'        .Parameters.Append .CreateParameter("p_VOIP_Type", adVarChar, adParamInput, 4000, str_VOIP_Type)
'        .Parameters.Append .CreateParameter("p_Dyna_IP_Num", adVarChar, adParamInput, 4000, str_Dyna_IP_Num)
'        .Parameters.Append .CreateParameter("p_Static_IP_Num", adVarChar, adParamInput, 4000, str_Static_IP_Num)
'        .Parameters.Append .CreateParameter("p_CPE_MAC", adVarChar, adParamInput, 4000, str_CPE_MAC)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_09"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_09 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_09"
'    Call_IEL_09 = False
'End Function
'
'Public Function Call_IEL_10(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByVal str_MTA_IP_Address As String, _
'                                                    Optional ByVal str_HFC_Node As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_10 As New ADODB.Command
'
'    With cmd_Call_IEL_10
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_MTA_IP_Address", adVarChar, adParamInput, 4000, str_MTA_IP_Address)
'        .Parameters.Append .CreateParameter("p_HFC_Node", adVarChar, adParamInput, 4000, str_HFC_Node)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_10"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_10 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_10"
'    Call_IEL_10 = False
'End Function
'
'Public Function Call_IEL_11(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByVal str_Modem_MAC_New As String, _
'                                                    Optional ByVal str_MTA_MAC As String, _
'                                                    Optional ByVal str_MTA_MAC_New As String, _
'                                                    Optional ByVal str_CustID As String, _
'                                                    Optional ByVal str_Equip_Type As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_11 As New ADODB.Command
'
'    With cmd_Call_IEL_11
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_Modem_MAC_New", adVarChar, adParamInput, 4000, str_Modem_MAC_New)
'        .Parameters.Append .CreateParameter("p_MTA_MAC", adVarChar, adParamInput, 4000, str_MTA_MAC)
'        .Parameters.Append .CreateParameter("p_MTA_MAC_New", adVarChar, adParamInput, 4000, str_MTA_MAC_New)
'        .Parameters.Append .CreateParameter("p_CustID", adVarChar, adParamInput, 4000, str_CustID)
'        .Parameters.Append .CreateParameter("p_Equip_Type", adVarChar, adParamInput, 4000, str_Equip_Type)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_11"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_11 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_11"
'    Call_IEL_11 = False
'End Function
'
'Public Function Call_IEL_12(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_12 As New ADODB.Command
'
'    With cmd_Call_IEL_12
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_12"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_12 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_12"
'    Call_IEL_12 = False
'End Function
'
'Public Function Call_IEL_13(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByVal str_MTA_MAC As String, _
'                                                    Optional ByVal str_CustID As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_13 As New ADODB.Command
'
'    With cmd_Call_IEL_13
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_MTA_MAC", adVarChar, adParamInput, 4000, str_MTA_MAC)
'        .Parameters.Append .CreateParameter("p_CustID", adVarChar, adParamInput, 4000, str_CustID)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_13"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_13 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_13"
'    Call_IEL_13 = False
'End Function
'
'Public Function Call_IEL_14(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByRef str_Client_Class As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_14 As New ADODB.Command
'
'    With cmd_Call_IEL_14
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Client_Class", adVarChar, adParamOutput, 4000, str_Client_Class)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_14"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         str_Client_Class = .Parameters("p_Client_Class").Value & ""
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_14 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_14"
'    Call_IEL_14 = False
'End Function
'
'Public Function Call_IEL_15(Optional ByVal str_Service_Provider As String, _
'                                                    Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByVal str_Msg_ID_2 As String, _
'                                                    Optional ByVal str_Msg_ID_3 As String, _
'                                                    Optional ByVal str_Modem_MAC As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, _
'                                                    Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_15 As New ADODB.Command
'
'    With cmd_Call_IEL_15
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_Msg_ID_2", adVarChar, adParamInput, 4000, str_Msg_ID_2)
'        .Parameters.Append .CreateParameter("p_Msg_ID_3", adVarChar, adParamInput, 4000, str_Msg_ID_3)
'        .Parameters.Append .CreateParameter("p_Modem_MAC", adVarChar, adParamInput, 4000, str_Modem_MAC)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_15"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_15 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_15"
'    Call_IEL_15 = False
'End Function
'
'Public Function Call_IEL_16(Optional ByVal str_Service_Provider As String, Optional ByVal str_Msg_ID_1 As String, _
'                                                    Optional ByRef str_IP_Address As String, Optional ByRef str_MAC_Address As String, _
'                                                    Optional ByRef str_Lease_Time As String, Optional ByRef str_Interface As String, _
'                                                    Optional ByRef str_Address_State As String, Optional ByRef str_IP_Lease_Start_Time As String, _
'                                                    Optional ByRef str_IP_Lease_Init_Start_Time As String, Optional ByRef str_IP_Lease_Protocol As String, _
'                                                    Optional ByRef str_IP_Lease_Remote_ID As String, Optional ByRef str_Subnet_Mask As String, _
'                                                    Optional ByRef str_Domain_Server As String, Optional ByRef str_Host_Name As String, _
'                                                    Optional ByRef str_T_FTP_Server_Name As String, Optional ByRef str_Boot_File_Name As String, _
'                                                    Optional ByRef str_Gateways As String, Optional ByRef str_Static_Address_Name As String, _
'                                                    Optional ByRef str_Static_Address_Desc As String, Optional ByRef str_Hardware_Address As String, _
'                                                    Optional ByRef str_Hardware_Address_Type As String, Optional ByRef str_Identify_Client As String, _
'                                                    Optional ByRef str_Static_IP As String, Optional ByRef str_Template As String, _
'                                                    Optional ByRef lng_Msg_ID As Long, Optional ByRef lng_Return As Long) As Boolean
'  On Error GoTo ChkErr
'
'    Dim cmd_Call_IEL_16 As New ADODB.Command
'
'    With cmd_Call_IEL_16
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue, 2000, lng_Return)
'        .Parameters.Append .CreateParameter("p_Service_Provider", adVarChar, adParamInput, 4000, str_Service_Provider)
'        .Parameters.Append .CreateParameter("p_Msg_ID_1", adVarChar, adParamInput, 4000, str_Msg_ID_1)
'        .Parameters.Append .CreateParameter("p_IP_Address", adVarChar, adParamOutput, 4000, str_IP_Address)
'        .Parameters.Append .CreateParameter("p_MAC_Address", adVarChar, adParamOutput, 4000, str_MAC_Address)
'        .Parameters.Append .CreateParameter("p_Lease_Time", adVarChar, adParamOutput, 4000, str_Lease_Time)
'        .Parameters.Append .CreateParameter("p_Interface", adVarChar, adParamOutput, 4000, str_Interface)
'        .Parameters.Append .CreateParameter("p_Address_State", adVarChar, adParamOutput, 4000, str_Address_State)
'        .Parameters.Append .CreateParameter("p_IP_Lease_Start_Time", adVarChar, adParamOutput, 4000, str_IP_Lease_Start_Time)
'        .Parameters.Append .CreateParameter("p_IP_Lease_Init_Start_Time", adVarChar, adParamOutput, 4000, str_IP_Lease_Init_Start_Time)
'        .Parameters.Append .CreateParameter("p_IP_Lease_Protocol", adVarChar, adParamOutput, 4000, str_IP_Lease_Protocol)
'        .Parameters.Append .CreateParameter("p_IP_Lease_Remote_ID", adVarChar, adParamOutput, 4000, str_IP_Lease_Remote_ID)
'        .Parameters.Append .CreateParameter("p_Subnet_Mask", adVarChar, adParamOutput, 4000, str_Subnet_Mask)
'        .Parameters.Append .CreateParameter("p_Domain_Server", adVarChar, adParamOutput, 4000, str_Domain_Server)
'        .Parameters.Append .CreateParameter("p_Host_Name", adVarChar, adParamOutput, 4000, str_Host_Name)
'        .Parameters.Append .CreateParameter("p_T_FTP_Server_Name", adVarChar, adParamOutput, 4000, str_T_FTP_Server_Name)
'        .Parameters.Append .CreateParameter("p_Boot_File_Name", adVarChar, adParamOutput, 4000, str_Boot_File_Name)
'        .Parameters.Append .CreateParameter("p_Gateways", adVarChar, adParamOutput, 4000, str_Gateways)
'        .Parameters.Append .CreateParameter("p_Static_Address_Name", adVarChar, adParamOutput, 4000, str_Static_Address_Name)
'        .Parameters.Append .CreateParameter("p_Static_Address_Desc", adVarChar, adParamOutput, 4000, str_Static_Address_Desc)
'        .Parameters.Append .CreateParameter("p_Hardware_Address", adVarChar, adParamOutput, 4000, str_Hardware_Address)
'        .Parameters.Append .CreateParameter("p_Hardware_Address_Type", adVarChar, adParamOutput, 4000, str_Hardware_Address_Type)
'        .Parameters.Append .CreateParameter("p_Identify_Client", adVarChar, adParamOutput, 4000, str_Identify_Client)
'        .Parameters.Append .CreateParameter("p_Static_IP", adVarChar, adParamOutput, 4000, str_Static_IP)
'        .Parameters.Append .CreateParameter("p_Template", adVarChar, adParamOutput, 4000, str_Template)
'        .Parameters.Append .CreateParameter("p_Msg_ID", adVarNumeric, adParamOutput, 2000, lng_Msg_ID)
'        .CommandType = adCmdStoredProc
'        .CommandText = "SF_Call_IEL_16"
'         Set .ActiveConnection = gcnGi
'        .Execute
'         str_IP_Address = .Parameters("p_IP_Address").Value & ""
'         str_MAC_Address = .Parameters("p_MAC_Address").Value & ""
'         str_Lease_Time = .Parameters("p_Lease_Time").Value & ""
'         str_Interface = .Parameters("p_Interface").Value & ""
'         str_Address_State = .Parameters("p_Address_State").Value & ""
'         str_IP_Lease_Start_Time = .Parameters("p_IP_Lease_Start_Time").Value & ""
'         str_IP_Lease_Init_Start_Time = .Parameters("p_IP_Lease_Init_Start_Time").Value & ""
'         str_IP_Lease_Protocol = .Parameters("p_IP_Lease_Protocol").Value & ""
'         str_IP_Lease_Remote_ID = .Parameters("p_IP_Lease_Remote_ID").Value & ""
'         str_Subnet_Mask = .Parameters("p_Subnet_Mask").Value & ""
'         str_Domain_Server = .Parameters("p_Domain_Server").Value & ""
'         str_Host_Name = .Parameters("p_Host_Name").Value & ""
'         str_T_FTP_Server_Name = .Parameters("p_T_FTP_Server_Name").Value & ""
'         str_Boot_File_Name = .Parameters("p_Boot_File_Name").Value & ""
'         str_Gateways = .Parameters("p_Gateways").Value & ""
'         str_Static_Address_Name = .Parameters("p_Static_Address_Name").Value & ""
'         str_Static_Address_Desc = .Parameters("p_Static_Address_Desc").Value & ""
'         str_Hardware_Address = .Parameters("p_Hardware_Address").Value & ""
'         str_Hardware_Address_Type = .Parameters("p_Hardware_Address_Type").Value & ""
'         str_Identify_Client = .Parameters("p_Identify_Client").Value & ""
'         str_Static_IP = .Parameters("p_Static_IP").Value & ""
'         str_Template = .Parameters("p_Template").Value & ""
'         lng_Msg_ID = Val(.Parameters("p_Msg_ID").Value & "")
'         lng_Return = Val(.Parameters("RETURN_VALUE").Value & "")
'    End With
'
'    Call_IEL_16 = (lng_Return = 0)
'
'  Exit Function
'ChkErr:
'    ErrSub "mod_CallSF", "Call_IEL_16"
'    Call_IEL_16 = False
'End Function

'Private Sub ShowMsg()
'  On Error GoTo ChkErr
'    strErrMsg = Empty
'    PgbComplete
'    If lngReturn = 0 Then
'        Select Case UCase(strStatus)
'                    Case "S"
'                                MsgBox "[ " & strCmdName & " ] 控制訊號已送出完成 !! ", vbInformation, gimsgPrompt
'                    Case "E"
'                                If lngReturn <> 0 Then
'                                    strErrMsg = GetErrMsg(CStr(lngReturn))
'                                Else
'                                    strErrMsg = "錯誤碼:0 , 未知錯誤 !"
'                                End If
'                                MsgBox "[ " & strCmdName & " ] 失敗 !! " & strCrlf & "錯誤原因 : " & vbCrLf & strErrMsg, vbInformation, gimsgPrompt
'                    Case "W"
'                                MsgBox "[ " & strCmdName & " ] 失敗 !! " & strCrlf & "錯誤原因 : 作業逾時 ( Time out ) !", vbInformation, gimsgPrompt
''                    Case "GT"
''                                MsgBox "[ " & strCmdName & " ] 失敗 !! " & strCrlf & "錯誤原因 : Gateway 與 IP Command time out !", vbInformation, gimsgPrompt
'        End Select
'    Else
'        Select Case lngReturn
'                    Case Is > 0
'                            If lngReturn <> 0 Then
'                                strErrMsg = GetErrMsg(strFailureCode)
'                            Else
'                                strErrMsg = GetErrMsg(CStr(lngReturn))
'                            End If
'                            MsgBox "[ " & strCmdName & " ] 失敗 !! " & strCrlf & "錯誤原因 : " & vbCrLf & strErrMsg, vbInformation, gimsgPrompt
'                    Case Is < 0
'                            MsgBox "[ " & strCmdName & " ] 失敗 !! " & strCrlf & "錯誤原因 : " & lngReturn, vbInformation, gimsgPrompt
'        End Select
'    End If
'  Exit Sub
'ChkErr:
'    ErrSub FormName, "subShowMessage"
'End Sub
