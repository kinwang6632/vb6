Attribute VB_Name = "mduControlPanelWeb"
Option Explicit

Private Const blnDebug = False

Public gCompFilterStr As String
Public gCompCode As String
Public gServiceType As String
Public gDefaultComp As String
Public objActForm As Form
Public gLogInID As String
Public gstrUserPremission As String
Public lngSeqNo As Long
Public strCMDSeqNo As String

'Public strCN As String
'Public gVerType As Integer
'Public gTestMode As Boolean
'Public gcnGi2 As New ADODB.Connection
'Dim objCmd As New CS_Web.clsEMCcommand

' 呼叫 EMC CM CP Web Command ( Call Web Service )
Public Function Call_CMCP_WebCmd(strID_Code As String, _
                                                            Optional strCompCode As String = "", _
                                                            Optional strCustID As String = "", _
                                                            Optional strCM_MAC As String = "", _
                                                            Optional strEMTA_MAC As String = "", Optional strEMTA_Port As String = "", _
                                                            Optional strCP_No As String = "", _
                                                            Optional strClass_ID As String = "", _
                                                            Optional strPC_IP_NO As String = "", _
                                                            Optional strPC_IP As String = "", _
                                                            Optional strOper_Type As String = "", _
                                                            Optional strStart_Date_Time As String = "", Optional strEnd_Date_Time As String = "", _
                                                            Optional lngADDRNO As Long = 0, _
                                                            Optional ByVal strDialAccount As String = "", Optional ByVal strDialPassword As String = "", _
                                                            Optional ByVal strTel1 As String = "", Optional ByVal strTel2 As String = "", _
                                                            Optional ByVal strCustName As String, _
                                                            Optional ByVal FinTime As String, _
                                                            Optional ByRef strCMDSeqNo As String, _
                                                            Optional ByVal strCmdType As String, _
                                                            Optional ByVal strID As String, _
                                                            Optional ByRef strResult As String = "", Optional ByRef strFaultReason As String = "", _
                                                            Optional ByRef strErrMsg As String = "", _
                                                            Optional ByRef objResult As Object, _
                                                            Optional ByVal strSNO As String = "", Optional ByVal strFaciSeqNo As String = "", _
                                                            Optional ByVal strNodeNo As String = "", Optional ByVal strCircuitNo As String = "") _
                                                            As Boolean
  On Error GoTo ChkErr
  
    Dim objWebCmd As Object
    Set objWebCmd = CreateObject("CS_Web.clsEMCcommand")
    If strCompCode = 0 Then strCompCode = gCompCode
    With objWebCmd
            Call_CMCP_WebCmd = .Web_Command(gcnGi.ConnectionString, _
                                                                            strCompCode, _
                                                                            GetOwner, _
                                                                            strID_Code, _
                                                                            garyGi(0), _
                                                                            strCompCode, _
                                                                            strCustID, _
                                                                            strCM_MAC, _
                                                                            strEMTA_MAC, strEMTA_Port, _
                                                                            strCP_No, _
                                                                            strClass_ID, _
                                                                            strPC_IP_NO, strPC_IP, _
                                                                            strOper_Type, _
                                                                            strStart_Date_Time, strEnd_Date_Time, _
                                                                            lngADDRNO, _
                                                                            strDialAccount, strDialPassword, _
                                                                            strTel1, strTel2, _
                                                                            strCustName, FinTime, _
                                                                            strCMDSeqNo, strCmdType, strID, _
                                                                            strResult, strFaultReason, _
                                                                            objResult, frmSO1623A.pbr1, _
                                                                            strSNO, strFaciSeqNo, _
                                                                            strNodeNo, strCircuitNo)
            strErrMsg = .GetErrMsg
    End With
    Set objWebCmd = Nothing
  Exit Function
ChkErr:
    ErrSub "Sys_Lib", "Call_CMCP_WebCmd"
End Function

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

'(1) CM 開 / 關/停/ 換/變更速率
Public Function OpenCM(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                        ByVal lngInstAddrNo As Long, ByVal strClassID As String, _
                                        ByVal strDynaIP As String, _
                                        ByRef strResult As String, ByRef strFaultReason As String, _
                                        ByRef strErrMsg As String, ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    Dim strGroupID As String, strCTIP As String
    If Not Call_CMCP_WebCmd("02", rsData("CompCode") & "", rsData("CustId") & "", _
                                            rsData("FaciSNo") & "", IIf(frmSO1623A.IsCMCP, Get_EMTA_Mac(rsData("FaciSNo") & ""), ""), , , _
                                            strClassID, strDynaIP, , 1, , , lngInstAddrNo, _
                                            , , , , , , strCMDSeqNo, , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
'        On Error Resume Next
'        strGroupID = rsResult("GroupID") & ""
        
        strCTIP = rsResult("CTIP") & ""
        If strCTIP <> Empty Then
            IPwarning rsData("IPAddress") & "", strCTIP
            rsData("IPAddress") = strCTIP
            rsData.Update
        End If
        OpenCM = True
    End If

    ShowMsg "(1) CM 開 / 關/停/ 換/變更速率", strResult, strFaultReason, strErrMsg

  Exit Function
ChkErr:
    ErrSub FormName, "OpenCM"
End Function

'(1) CM 開 / 關/停/ 換/變更速率
Public Function CloseCM(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                        ByVal lngInstAddrNo As Long, ByVal strClassID As String, _
                                        ByVal strDynaIP As String, _
                                        ByRef strResult As String, ByRef strFaultReason As String, _
                                        ByRef strErrMsg As String, ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    Dim strGroupID As String, strCTIP As String
    If Not Call_CMCP_WebCmd("02", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNo") & "", Get_EMTA_Mac(rsData("FaciSNo") & ""), , , _
                                                strClassID, strDynaIP, , 2, , , lngInstAddrNo, _
                                                , , , , , , strCMDSeqNo, , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
'        strGroupID = rsResult("GroupID") & ""
        strCTIP = rsResult("CTIP") & ""
        If strCTIP <> Empty Then
            IPwarning rsData("IPAddress") & "", strCTIP
            rsData("IPAddress") = strCTIP
            rsData.Update
        End If
        CloseCM = True
    End If
    
    ShowMsg "(1) CM 開 / 關/停/ 換/變更速率", strResult, strFaultReason, strErrMsg

  Exit Function
ChkErr:
    ErrSub FormName, "CloseCM"
End Function

'(1) CM 開 / 關/停/ 換/變更速率
Public Function StopCM(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                        ByVal lngInstAddrNo As Long, ByVal strClassID As String, _
                                        ByVal strDynaIP As String, _
                                        ByRef strResult As String, ByRef strFaultReason As String, _
                                        ByRef strErrMsg As String, ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    Dim strGroupID As String, strCTIP As String
    If Not Call_CMCP_WebCmd("02", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNo") & "", Get_EMTA_Mac(rsData("FaciSNo") & ""), , , _
                                                strClassID, strDynaIP, , 1, , , lngInstAddrNo, _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
'        strGroupID = rsResult("GroupID") & ""
        strCTIP = rsResult("CTIP") & ""
        If strCTIP <> Empty Then
            IPwarning rsData("IPAddress") & "", strCTIP
            rsData("IPAddress") = strCTIP
            rsData.Update
        End If
        StopCM = True
    End If
    
    ShowMsg "(1) CM 開 / 關/停/ 換/變更速率", strResult, strFaultReason, strErrMsg

  Exit Function
ChkErr:
    ErrSub FormName, "StopCM"
End Function

Private Sub IPwarning(ByVal strOriIP As String, ByVal strNewIP As String)
  On Error Resume Next
    If Len(strOriIP) > 0 Then
        If strOriIP <> strNewIP Then MsgBox "所取得之IP [ " & strNewIP & " ] 與原取得IP [ " & strOriIP & " ] 不同 ， 請確認 EMTA 是否可正常運作 !", vbInformation, "訊息"
    End If
End Sub

'(1) CM 開 / 關/停/ 換/變更速率
Public Function ChangeCM(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                            ByVal lngInstAddrNo As Long, ByVal strClassID As String, _
                                            ByVal strDynaIP As String, _
                                            ByRef strResult As String, ByRef strFaultReason As String, _
                                            ByRef strErrMsg As String, ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    Dim strGroupID As String, strCTIP As String
    If Not Call_CMCP_WebCmd("02", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNo") & "", Get_EMTA_Mac(rsData("FaciSNo") & ""), , , _
                                                strClassID, strDynaIP, , 1, , , lngInstAddrNo, _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
'        strGroupID = rsResult("GroupID") & ""
        strCTIP = rsResult("CTIP") & ""
        If strCTIP <> Empty Then
            IPwarning rsData("IPAddress") & "", strCTIP
            rsData("IPAddress") = strCTIP
            rsData.Update
        End If
        ChangeCM = True
    End If

    ShowMsg "(1) CM 開 / 關/停/ 換/變更速率", strResult, strFaultReason, strErrMsg

  Exit Function
ChkErr:
    ErrSub FormName, "ChangeCM"
End Function

'(1) CM 開 / 關/停/ 換/變更速率
Public Function ChangeCMRate(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                                    ByVal lngInstAddrNo As Long, ByVal strClassID As String, _
                                                    ByVal strDynaIP As String, _
                                                    ByRef strResult As String, ByRef strFaultReason As String, _
                                                    ByRef strErrMsg As String, ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    Dim strGroupID As String, strCTIP As String
    If Not Call_CMCP_WebCmd("02", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNo") & "", Get_EMTA_Mac(rsData("FaciSNo") & ""), , , _
                                                strClassID, strDynaIP, , 1, , , lngInstAddrNo, _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
'        strGroupID = rsResult("GroupID") & ""
        strCTIP = rsResult("CTIP") & ""
        If strCTIP <> Empty Then
            IPwarning rsData("IPAddress") & "", strCTIP
            rsData("IPAddress") = strCTIP
            rsData.Update
        End If
        ChangeCMRate = True
    End If
    
    ShowMsg "(1) CM 開 / 關/停/ 換/變更速率", strResult, strFaultReason, strErrMsg

  Exit Function
ChkErr:
    ErrSub FormName, "ChangeCMRate"
End Function

'(2) CP IP 申請及退租
Public Function CPInst(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                    ByVal lngInstAddrNo As Long, ByRef strResult As String, _
                                    ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                    ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    CPInst = CPIP("1", rsData, strPCMac, lngInstAddrNo, strResult, strFaultReason, strErrMsg, rsResult)
  Exit Function
ChkErr:
    ErrSub FormName, "CPInst"
End Function

Private Function CPIP(strOperType As String, _
                                    rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                    ByVal lngInstAddrNo As Long, ByRef strResult As String, _
                                    ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                    ByRef rsResult As ADODB.Recordset) As Boolean
    
    Dim rsTmp As New ADODB.Recordset
    Dim strCTIP As String
    
    Dim strQry As String
    
    strQry = "SELECT PORTNO,FACISNO FROM " & GetOwner & "SO004" & _
                    " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                    " AND COMPCODE=" & rsData("CompCode") & _
                    " AND CUSTID=" & rsData("CustID") & _
                    " AND SERVICETYPE='P'" & _
                    " AND INSTDATE IS NOT NULL" & _
                    " AND PRDATE IS NULL"
    ' Liga :
    '   2．CP申裝[27]：少3個參數：CMMAC, EMTAMAC, EMTAPORT。CPNO傳值有誤，應傳P服務的FACISNO才對，現是傳REFACISNO
    
    CPIP = False
    
    If GetRS(rsTmp, strQry, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then
        With rsTmp
            If .RecordCount > 0 Then
                While Not .EOF
                    If Not Call_CMCP_WebCmd("27", rsData("CompCode") & "", rsData("CustId") & "", _
                                                                rsData("FaciSNo") & "", Get_EMTA_Mac(rsData("FaciSNo") & ""), _
                                                                rsTmp("PortNo") & "", rsTmp("FaciSNO") & "", , , , strOperType, , , lngInstAddrNo, _
                                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
                    Else
                        strCTIP = rsResult("CTIP") & ""
                        If strCTIP <> Empty Then
                            IPwarning rsData("IPAddress") & "", strCTIP
                            rsData("IPAddress") = strCTIP
                            rsData.Update
                        End If
                        CPIP = True
                    End If
                    .MoveNext
                Wend
            Else
                ' 找不到 P 服務資料
                strFaultReason = "此設備無安裝 CP !"
'                MsgBox "此設備無安裝 CP !", vbInformation, "訊息"
            End If
        End With
    End If
    
    ShowMsg "(2) CP IP 申請及退租", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "CPIP"
End Function

'(2) CP IP 申請及退租
Public Function CPPR(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                    ByVal lngInstAddrNo As Long, ByRef strResult As String, _
                                    ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                    ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    ' Liga :    3．CP退租[27]：OPERTYPE有誤，應傳值2才對。
    CPPR = CPIP("2", rsData, strPCMac, lngInstAddrNo, strResult, strFaultReason, strErrMsg, rsResult)
  Exit Function
ChkErr:
    ErrSub FormName, "CPPR"
End Function

'(3) Reset CM
Public Function ResetCM(rsData As ADODB.Recordset, ByRef strResult As String, _
                                        ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                        ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("12", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNO") & "", , , , , , , , , , , _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
        ResetCM = True
    End If
    
    ShowMsg "(3) Reset CM", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "ResetCM"
End Function

'(4) Ping CM
Public Function PingCM(rsData As ADODB.Recordset, ByRef strResult As String, _
                                        ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                        ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("20", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNO") & "", , , , , , , , , , , _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
        PingCM = True
    End If
    
    ShowMsg "(4) Ping CM", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "PingCM"
End Function

''(5) 釋放 動態IP
'Public Function ReleaseDyIP(rsData As ADODB.Recordset, ByRef strResult As String, _
'                                                ByRef strFaultReason As String, ByRef strErrMsg As String, _
'                                                ByRef rsResult As Object) As Boolean
'  On Error GoTo ChkErr
'    If Not Call_CMCP_WebCmd("25", rsData("CompCode"), rsData("CustId"), _
'                                                rsData("FaciSNO"), , , , , , , , , , , _
'                                                strResult, strFaultReason, strErrMsg, rsResult) Then
'        Exit Function
'    End If
'
'    ReleaseDyIP = True
'  Exit Function
'ChkErr:
'    ErrSub FormName, "ReleaseDyIP"
'End Function

''(7) ACS帳號密碼相關功能
'Public Function ACSAccPWS(rsData As ADODB.Recordset, ByRef strResult As String, _
'                                            ByRef strFaultReason As String, ByRef strErrMsg As String, _
'                                            ByRef rsResult As Object) As Boolean
'  On Error GoTo ChkErr
'    If Not Call_CMCP_WebCmd("25", rsData("CompCode"), rsData("CustId"), _
'                                                rsData("FaciSNO"), , , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
'        Exit Function
'    End If
'
'    ACSAccPWS = True
'  Exit Function
'ChkErr:
'    ErrSub FormName, "ACSAccPWS"
'End Function

'(7) Clear CPE IP
Public Function ClearCPEIP(rsData As ADODB.Recordset, ByRef strResult As String, _
                                            ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                            ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("25", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNO") & "", , , , , , , , , , , _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
        ClearCPEIP = True
    End If
    
    ShowMsg "(5) 釋放動態IP", strResult, strFaultReason, strErrMsg

  Exit Function
ChkErr:
    ErrSub FormName, "ClearCPEIP"
End Function

'(6) CM 狀態查詢
Public Function QueryCMStatus(rsData As ADODB.Recordset, ByVal lngInstAddrNo As Long, _
                                                    ByRef strResult As String, ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                                    ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("13", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNO") & "", , , , , , , , , , lngInstAddrNo, _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult, _
                                                rsData("SNO") & "", rsData("SeqNo") & "") Then
    Else
        Dim strPCIP As String
        Dim varAry As Variant
        Dim varElement As Variant
        strPCIP = rsResult("PCIP") & ""
        If Len(strPCIP) > 0 Then
            gcnGi.Execute "DELETE FROM " & GetOwner & "SO004C" & _
                                    " WHERE CUSTID=" & rsData("CustID") & _
                                    " AND FACISNO='" & rsData("FaciSno") & "'" & _
                                    " AND SEQNO='" & rsData("SeqNo") & "'"
                                    
            If InStr(1, strPCIP, ",") Then
                varAry = Split(strPCIP, ",")
                For Each varElement In varAry
                    If varElement & "" <> Empty Then
                        gcnGi.Execute "INSERT INTO " & GetOwner & "SO004C" & _
                                                " (CUSTID,FACISNO,SEQNO,CPEMAC,IPAddress)" & _
                                                " VALUES (" & rsData("CustID") & _
                                                ",'" & rsData("FaciSno") & "'" & _
                                                ",'" & rsData("SeqNo") & "'" & _
                                                ",NULL,'" & varElement & "')"
                    End If
                Next
            Else
                gcnGi.Execute "INSERT INTO " & GetOwner & "SO004C" & _
                                        " (CUSTID,FACISNO,SEQNO,CPEMAC,IPAddress)" & _
                                        " VALUES (" & rsData("CustID") & _
                                        ",'" & rsData("FaciSno") & "'" & _
                                        ",'" & rsData("SeqNo") & "'" & _
                                        ",NULL,'" & rsResult("PCIP") & "')"
            End If
        End If
        ' 這個命令有 Return 一堆資料
        ' QueryResult
        ' FaultReason
        ' GroupID
        ' CMStatus
        ' CMIP
        ' CMUpFreq
        ' CMDownFreq
        ' CMRECPW
        ' CMTransPW
        ' CMSNR
        ' CMOnlineTimes
        ' CMInOctets
        ' CMOutOctets
        ' CMInErrors
        ' CMOutErrors
        ' PCIP
        ' CMTSRECPW
        ' CMTSRXSNR
        ' CMTSUPMOD
        QueryCMStatus = True
    End If
    
    ShowMsg "(6) CM 狀態查詢", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "QueryCMStatus"
End Function

'(8) 查詢 HFC 服務類別
Public Function QueryHFCServ(rsData As ADODB.Recordset, ByVal lngInstAddrNo As Long, _
                                                ByRef strResult As String, ByRef strFaultReason As String, _
                                                ByRef strErrMsg As String, ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    
    Dim strDualCable As String
    Dim strNodeNo As String
    Dim lngRcdAft As Long
    Dim strCNorNN As String ' CircuitNo or NodeNo
    
    If lngInstAddrNo = 0 Then
        MsgBox "地址編號 [AddrNo] 不正確!", vbInformation, "訊息"
        Exit Function
    End If
    If Not Call_CMCP_WebCmd("21", rsData("CompCode") & "", , , , , , , , , , , , _
                                                lngInstAddrNo, , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
'        MsgBox strErrMsg, vbInformation, "訊息"
    Else
        ' 這個命令有 Return 一堆資料
        ' OperResult
        ' FaultReason
        ' HFCType = SO014.DualCable
        ' NODEID = SO014.NodeNo
        
        If strResult = "1" Then
        
            ShowMsg "(8) 查詢 HFC 服務類別", strResult, strFaultReason, strErrMsg
        
            With rsResult
                    If .EOF Then .MoveFirst
                    strDualCable = .Fields("HFCTYPE") & ""
                    strNodeNo = .Fields("NODEID") & ""
            End With
            
            ' SO041.EMTAIPTYPE ( 0=網路編號 , 1=Node )
            If GetSystemParaItem("EMTAIPTYPE", gCompCode, gServiceType, "SO041", , True) = 0 Then
                strCNorNN = "CIRCUITNO"
            Else
                strCNorNN = "NODENO"
            End If
            
'                If strDualCable <> Empty Then strDualCable = Choose(Left(strDualCable, 1), "0", "1", "2", "3")
            ' DualCable (HFCTYPE) 2 , 3 的部份 Jim 說忽略掉, 所以不處理
            If strDualCable <> Empty Then strDualCable = Choose(Left(strDualCable, 1), "0", "1") & ""
            
            lngRcdAft = 0
            
            If strDualCable <> "" And strNodeNo <> "" Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO014" & _
                                        " SET DUALCABLE=" & strDualCable & "," & _
                                        " " & strCNorNN & "='" & strNodeNo & "'" & _
                                        " WHERE ADDRNO=" & lngInstAddrNo & _
                                        " AND COMPCODE=" & gCompCode, lngRcdAft
                                        
            ElseIf strDualCable <> "" Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO014" & _
                                        " SET DUALCABLE=" & strDualCable & _
                                        " WHERE ADDRNO=" & lngInstAddrNo & _
                                        " AND COMPCODE=" & gCompCode, lngRcdAft
                                        
            ElseIf strNodeNo <> "" Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO014" & _
                                        " SET " & strCNorNN & "='" & strNodeNo & "'" & _
                                        " WHERE ADDRNO=" & lngInstAddrNo & _
                                        " AND COMPCODE=" & gCompCode, lngRcdAft
            Else
                MsgBox "回傳 [HFCType] , [NODEID] 資料為空 ! 地址資料未更新!", vbInformation, "訊息"
            End If
                                    
            If lngRcdAft > 0 Then
                MsgBox "地址資料 [HFCType] , [NODEID] 更新完成!", vbInformation, "訊息"
            Else
                MsgBox "查無地址資料 , 無法更新!", vbInformation, "訊息"
            End If
            
'                Debug.Print rsResult.GetString

            ' QQ ? Schema 裡註記 NodeNo 由網路編號 (CircuitNo) 前三碼存入
            
            '於客戶主檔[SO1100B]的畫面增加『網工資料』查詢的按鍵；
            '●使用者點選此按鍵時請依據　此份文件　CC&B作業流程與SPM　介接流程圖_Jim_20050805.XLS
            '  的這個頁籤　“EMC　WEB　Command　matrix　”至
            '  EMC網管的介接的程式傳送第8個指令
            '  “查詢　HFC　服務類別　HFCTypeQuery　21”
            '●並且依據此客戶的地址資料取得網路編號或NODENO後,　請存入此客戶裝機地址的正確欄位
            '  (SO001.INSTADDRNO=SO014.ADDRNO,　將取回來的值回填至SO014.Circuitno　或　so014.Nodeno)

        Else
'            MsgBox "執行結果 : " & strResult & vbCrLf & "失敗原因 : " & strFaultReason, vbInformation, "訊息"
            ShowMsg "(8) 查詢 HFC 服務類別", strResult, strFaultReason, strErrMsg
        End If
        
        QueryHFCServ = True
    End If
    
  Exit Function
ChkErr:
    ErrSub FormName, "QueryHFCServ"
End Function

'(9) CM 連線紀錄查詢
Public Function QueryCMHistory(rsData As ADODB.Recordset, ByVal strStartDate As String, _
                                                    ByVal strEndDate As String, ByRef strResult As String, _
                                                    ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                                    ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("22", rsData("CompCode") & "", rsData("CustId") & "", rsData("FaciSNo") & "", _
                                                , , , , , , , strStartDate, strEndDate, , _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
        ' 這個命令有 Return 一堆資料
        ' QueryResult
        ' FaultReason
        ' CMIP
        ' PCIP = SO004C按照IP數與MAC等關聯資訊，按結果作異動或新增或刪除
        ' OffLineTime
        ' OnLineTime
        QueryCMHistory = True
    End If
    
    ShowMsg "(9) CM 連線紀錄查詢", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "QueryCMHistory"
End Function

'(10) CM 連線品質查詢
Public Function QueryCMQuality(rsData As ADODB.Recordset, ByVal strStartDate As String, _
                                                    ByVal strEndDate As String, ByRef strResult As String, _
                                                    ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                                    ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("23", rsData("CompCode") & "", rsData("CustId") & "", rsData("FaciSNo") & "", _
                                                , , , , , , , strStartDate, strEndDate, , _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
        ' 這個命令有 Return 一堆資料
        ' QueryResult
        ' FaultReason
        ' CMRECPW
        ' CMTransPW
        ' CMSNR
        ' CMInOctets
        ' CMOutOctets
        ' CMInErrors
        ' CMOutErrors
        ' CMTSRECPW
        ' CMTSRXSNR
        ' DetctTime
        QueryCMQuality = True
    End If
    
    ShowMsg "(10) CM 連線品質查詢", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "QueryCMQuality"
End Function

'(11)PCIP連線記錄查詢
Public Function QueryPCIPHistory(rsData As ADODB.Recordset, ByVal strPCIP As String, _
                                                        ByVal strStartDate As String, ByVal strEndDate As String, _
                                                        ByRef strResult As String, ByRef strFaultReason As String, _
                                                        ByRef strErrMsg As String, ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("26", rsData("CompCode") & "", rsData("CustId") & "", _
                                                , , , , , , strPCIP, , strStartDate, strEndDate, , _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
        ' 這個命令有 Return 一堆資料
        ' QueryResult
        ' FaultReason
        ' PConlineTime
        ' PCofflineTime
        ' CMMAC
        QueryPCIPHistory = True
    End If
    
    ShowMsg "(11) PCIP連線記錄查詢", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "QueryPCIPHistory"
End Function

' 帳號啟用停用
Public Function Cmd_Acc_24(ByRef rsData As Object, strCMD As String, Optional strCMDSeqNo As String = "") As Boolean
  On Error GoTo ChkErr
    
    Dim strID As String
    
    'IDKindCode1 ID IDKindCode ID2
     
     With rsData
        
        If .Fields("IDKindCode1") & "" <> "" Then
            If GetRsValue("SELECT NVL(REFNO,0) FROM " & GetOwner & "CD070" & _
                                   " WHERE CODENO = " & .Fields("IDKindCode1")) = 1 Then
                strID = .Fields("ID") & ""
            End If
        End If
        
        If .Fields("IDKindCode") & "" <> "" Then
            If GetRsValue("SELECT NVL(REFNO,0) FROM " & GetOwner & "CD070" & _
                                    " WHERE CODENO = " & .Fields("IDKindCode")) = 1 Then
                strID = .Fields("ID2") & ""
            End If
        End If
     
     End With
    
    Cmd_Acc_24 = Call_CMCP_WebCmd("24", rsData("CompCode"), rsData("CustId"), _
                                                rsData("FaciSNo"), Get_EMTA_Mac(rsData("FaciSNo") & ""), _
                                                , , , , , strCMD, , , , rsData("DialAccount") & "", , , , , , strCMDSeqNo, strCMD, strID)
    
    If Cmd_Acc_24 Then MsgBox "已完成開通記錄 !", vbInformation, "訊息"



'    disable: 帳號停權
'    Enable: 帳號啟用
                                                    
'    系統台別 Cusso
'    客戶編號 Cusid
'    CM MAC  CM MAC
'    功能指令(開關)  OperType
'    撥接帳號 DialAccount
'    指令 log table 流水號   CmdSeqno
  
  Exit Function
ChkErr:
    ErrSub FormName, "Cmd_Acc_24"
End Function


Private Sub ShowMsg(strCMD As String, strResult As String, strFaultReason As String, strErrMsg As String)
    On Error GoTo ChkErr
        If strResult = "1" Then
            MsgBox "[ " & strCMD & " ] 控制訊號已送出完成 !! ", vbInformation, gimsgPrompt
        Else
            If strFaultReason = Empty Then
                MsgBox "[ " & strCMD & " ] 失敗 !! " & strCrlf & _
                                "錯誤原因 : " & strErrMsg, vbInformation, gimsgPrompt
            Else
                MsgBox "[ " & strCMD & " ] 失敗 !! " & strCrlf & _
                                "錯誤原因 : " & strFaultReason & strCrlf & _
                                strErrMsg, vbInformation, gimsgPrompt
            End If
        End If
    Exit Sub
ChkErr:
    ErrSub FormName, "subShowMessage"
End Sub



