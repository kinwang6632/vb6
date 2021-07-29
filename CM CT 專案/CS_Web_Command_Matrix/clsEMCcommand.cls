VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEMCcommand"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' 地址相關資料變數宣告
Private strZone As String
Private strLin As String
Private strSection As String
Private strLane As String
Private strAlley As String
Private strSubAlley As String
Private strNo1 As String
Private strNo2 As String

' Web Service 回傳相關變數宣告
Private strOperResult As String
Private strQueryResult As String
Private strFaultReason As String
Private strGroupID As String
Private strCMStatus As String
Private strR_CMIP As String
Private strR_CMUpFreq As String
Private strR_CMDownFreg As String
Private strR_CMRecPw As String
Private strR_CMTransPw As String
Private strR_CMSNR As String
Private strR_CMOnLineTimes As String
Private strR_CMInOctets As String
Private strR_CMOutOtets As String
Private strR_CMInErrors As String
Private strR_CMOutErrors As String
Private strR_PCIP As String
Private strR_CMTsRecPw As String
Private strR_CMTsRxSNR As String
Private strR_CMTSUpMod As String
Private strR_HFCType As String
Private strR_PCOffLineTime As String
Private strR_PCOnLineTime As String
Private strR_OffLineTime As String
Private strR_OnLineTime As String
Private strR_DetctTime As String
Private strR_CMMAC As String
Private strR_CTIP As String
Private strR_NODEID As String
'Private strExecEntry As String
'Private strExecDate As String
'Private strHostName As String

Private str_DialAccount As String
Private str_DialPassword As String
Private str_Tel1 As String
Private str_Tel2 As String
Private str_CustName As String
Private str_FinTime As String
Private str_CmdSeqNo As String
Private str_CmdType As String
Private str_ID As String
Private str_SNO As String
Private str_FaciSeqNo As String
Private str_NodeNo As String
Private str_CircuitNo As String

Private lngCMtimeout As Long
Private strCMowner As String
Private strCHcmdMsg As String

Private Const YMDHMSFmt As String = "TO_DATE(?,'YYYY/MM/DD HH24:MI:SS')"
Private Const YMDFmt As String = "TO_DATE(?,'YYYY/MM/DD')"

'Private Sub Class_Initialize()
'  On Error GoTo ChkErr
''    Set rsObj = CreateObject("ADODB.Recordset")
''    Licenses.Add "ADODB.Connection"
''    Licenses.Add "ADODB.Recordset"
'  Exit Sub
'ChkErr:
'    ErrHandle , "Class_Initialize"
'End Sub

'----------------------------------------------------------------------------------------
'|                                                           |    ID_CODE                            |
'----------------------------------------------------------------------------------------
'|    員工編號                                           |    WorkID                               |
'----------------------------------------------------------------------------------------
'|    系統台別                                           |    Cusso                                 |
'----------------------------------------------------------------------------------------
'|    客戶編號                                           |    Cusid                                  |
'----------------------------------------------------------------------------------------
'|    CM MAC                                           |    CM MAC                              |
'----------------------------------------------------------------------------------------
'|    EMTA MAC                                       |     EMTAMAC                           |
'----------------------------------------------------------------------------------------
'|    EMTA Port                                       |     EMTAPort                           |
'----------------------------------------------------------------------------------------
'|    CP電話號碼                                       |    CPNO                                  |
'----------------------------------------------------------------------------------------
'|    方案類別                                           |    ClassID                               |
'----------------------------------------------------------------------------------------
'|    動態PCIP數                                       |    PCIPNO  ( 原為PCIPNOD )   |
'----------------------------------------------------------------------------------------
'***|---靜態PCIP數---------------------------|----PCIPNOS---------------------|*** 拿掉了
'----------------------------------------------------------------------------------------
'***|---靜態PCIP MAC-----------------------|----PCMAC------------------------|*** 拿掉了
'----------------------------------------------------------------------------------------
'***|---客戶帳號------------------------------|----CUSACCO---------------------|*** 拿掉了
'----------------------------------------------------------------------------------------
'***|---客戶密碼------------------------------|----CUSPASSWORD--------------|*** 拿掉了
'----------------------------------------------------------------------------------------
'|    PCIP                                                |    PCIP                                   |
'----------------------------------------------------------------------------------------
'|    功能指令(開關停機,新增刪除修改查詢)    |    OperType                            |
'----------------------------------------------------------------------------------------
'|    起始時間                                           |    StartDateTime                     |
'----------------------------------------------------------------------------------------
'|    結束時間                                           |    EndDateTime                       |
'----------------------------------------------------------------------------------------
'|    地址資訊                                           |                                                |
'----------------------------------------------------------------------------------------
'|    鄉鎮市\村里莊埔路街\街道名稱              |    Zone                                   |
'----------------------------------------------------------------------------------------
'|    鄰                                                    |    Lin                                       |
'----------------------------------------------------------------------------------------
'|    段                                                    |    Section                                |
'----------------------------------------------------------------------------------------
'|    巷                                                    |    Lane                                    |
'----------------------------------------------------------------------------------------
'|    弄                                                    |    Alley                                    |
'----------------------------------------------------------------------------------------
'|    衖                                                    |    SubAlley                              |
'----------------------------------------------------------------------------------------
'|    號一                                                 |    NO1                                     |
'----------------------------------------------------------------------------------------
'|    號二                                                 |    NO2                                     |
'----------------------------------------------------------------------------------------
'|********************** 新增 ****************************|
'----------------------------------------------------------------------------------------
'|    撥接帳號                                           |    DialAccount                          |
'----------------------------------------------------------------------------------------
'|    撥接密碼                                           |    DialPassword                      |
'----------------------------------------------------------------------------------------
'|    電話一                                              |    Tel1                                     |
'----------------------------------------------------------------------------------------
'|    聯絡人電話                                        |    Tel2                                     |
'----------------------------------------------------------------------------------------
'|    申裝人姓名                                        |    CustName                            |
'----------------------------------------------------------------------------------------
'|    身份證字號                                        |    ID                                        |
'----------------------------------------------------------------------------------------
'|    完工時間                                           |    FinTime                                |
'----------------------------------------------------------------------------------------
'|    指令 log table 流水號                        |    CmdSeqNo                          |
'----------------------------------------------------------------------------------------

' 呼叫 EMC Web Command Matrix ---- 2005/08/01  By Hammer
' Last modify ---- 2005/09/02  By Hammer
Public Function Web_Command(ByVal strCN As String, _
                                                    ByVal strComp_Code As String, ByVal strOwner_Name As String, _
                                                    ByVal strID_Code As String, ByVal strWork_ID As String, _
                                                    ByVal strCust_SO As String, Optional ByVal strCustID As String = "", _
                                                    Optional ByVal strCM_MAC As String = "", _
                                                    Optional ByVal strEMTA_MAC As String = "", Optional ByVal strEMTA_Port As String = "", _
                                                    Optional ByVal strCP_No As String = "", _
                                                    Optional ByVal strClass_ID As String = "", _
                                                    Optional ByVal strPC_IP_NO As String = "", _
                                                    Optional ByVal strPC_IP As String = "", _
                                                    Optional ByVal strOper_Type As String = "", _
                                                    Optional ByVal strStart_Date_Time As String = "", Optional ByVal strEnd_Date_Time As String = "", _
                                                    Optional ByVal lngAddrNo As Long = 0, _
                                                    Optional ByVal strDialAccount As String = "", Optional ByVal strDialPassword As String = "", _
                                                    Optional ByVal strTel1 As String = "", Optional ByVal strTel2 As String = "", _
                                                    Optional ByVal strCustName As String = "", Optional ByVal FinTime As String = "", _
                                                    Optional ByRef strCmdSeqNo As String = "", Optional ByVal strCmdType As String = "", _
                                                    Optional ByVal strID As String = "", _
                                                    Optional ByRef strResult As String = "", Optional ByRef strFaultReason As String = "", _
                                                    Optional ByRef objResult As Object, _
                                                    Optional ByRef objPGB As Object = Nothing, _
                                                    Optional ByVal strSNO As String = "", Optional ByVal strFaciSeqNo As String = "", _
                                                    Optional ByVal strNodeNo As String = "", Optional ByVal strCircuitNo As String = "", _
                                                    Optional ByVal strPrcDate As String = "", _
                                                    Optional ByVal blnBatch As Boolean = False) As Boolean
  On Error GoTo ChkErr
  
    DoEvents
    
    Web_Command = False
    
    Dim objCN As Object
    
    Set objCN = CreateObject("ADODB.Connection")
    With objCN
            .CursorLocation = 3
            .Open strCN
    End With
    
    ' Check Oracle database connection status
    If objCN.State <= 0 Then
        strErrorMessage = "資料庫連線尚未開啟 !"
        Exit Function
    End If
    
    strOwner = strOwner_Name ' Set global variant "Owner"
    intCompCode = strComp_Code ' Set global variant "CompCode"
    
    str_DialAccount = strDialAccount
    str_DialPassword = strDialPassword
    str_Tel1 = strTel1
    str_Tel2 = strTel2
    str_CustName = strCustName
    str_FinTime = FinTime
    str_CmdSeqNo = strCmdSeqNo
    str_CmdType = strCmdType
    str_ID = strID
    
    str_SNO = strSNO
    str_FaciSeqNo = strFaciSeqNo
    str_NodeNo = strNodeNo
    str_CircuitNo = strCircuitNo
    
    ' 取得介接 Web Service 主機位置資訊
    On Error Resume Next
    strURL = objCN.Execute("SELECT CMGATEWAYSERVER FROM " & strOwner & "CD039" & _
                                            " WHERE CODENO=" & strComp_Code).GetString(2, 1, "", "", "") & ""
    If Err.Number <> 0 Then
        Err.Clear
        strErrorMessage = "[CD039] 公司別代碼檔 查無 [公司別] 為 " & strComp_Code & " 的資料 , 請確認 !"
        Exit Function
    Else
        strURL = RPxx(strURL)
    End If
    
    On Error GoTo ChkErr
    
    If strURL = "" Or strURL = vbCr Or strURL = vbCrLf Or strURL = Empty Then
        strErrorMessage = "[CD039] 公司別代碼檔 [介接主機位址資訊] 尚未設定, 請確認 !"
        Exit Function
    End If
    
    ' 檢核 CM CP 啟動介接參數 及 CMTimeOut , CMOwner(CM共用區Owner)
    On Error Resume Next
    
    Dim varSysPara As Variant
    Dim strQryPara As String
    
'   SO041.CMTimeOut
'   SO041.CMOwner
    
    strQryPara = "SELECT INTEGRATESPM,CMTIMEOUT,CMOWNER FROM " & strOwner & "SO041" & _
                                        " WHERE COMPCODE=" & strComp_Code
                                        
    varSysPara = objCN.Execute(strQryPara).GetRows(1, 1)
    
     ' Check INTEGRATESPM
    If Val(varSysPara(0, 0) & "") = 0 Then
        strErrorMessage = "[SO041] 系統業者參數檔 [啟動CMCP介接] 尚未開啟, 因此並未呼叫 Web Service 介接 !"
        Exit Function
    End If
    
     ' Check CMTimeOut
    lngCMtimeout = Val(varSysPara(1, 0) & "")
    If lngCMtimeout = 0 Then lngCMtimeout = 30 ' Default
    
     ' Check CMOwner
    strCMowner = varSysPara(2, 0) & ""
    If strCMowner = Empty Then
'        strErrorMessage = "[SO041] 系統業者參數檔 [CM共用區Owner] 尚未設定, 請確認 !"
        strCMowner = ""
    Else
        strCMowner = strCMowner & "."
    End If
    
'    If Val(objCN.Execute("SELECT INTEGRATECMCP FROM " & strOwner & "SO041" & _
'                                        " WHERE COMPCODE=" & strComp_Code).GetString(2, 1, "", "", "")) = 0 Then
'        strErrorMessage = "[SO041] 系統業者參數檔 [啟動CMCP介接] 尚未開啟, 因此並未呼叫 Web Service 介接 !"
'        Exit Function
'    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        strErrorMessage = "[SO041] 系統業者參數檔 查無 [公司別] 為 " & strComp_Code & " 的資料 , 請確認 !"
        Exit Function
    End If
    
    On Error GoTo ChkErr
    Dim varPara As Variant
    
    strErrorMessage = Empty
    
    Clear_Addr_Data_Var ' 清空地址資料相關變數

    Clear_Output_Para ' 清空 Web Service 回傳相關變數宣告
    
'    If strID_Code <> "24" Then strCmdSeqNo = Empty
    strCmdSeqNo = Empty
    If strCmdSeqNo = Empty Then str_CmdSeqNo = Get_Cmd_Seq_No(objCN, strPrcDate)
    strCmdSeqNo = str_CmdSeqNo
    
    Select Case strID_Code
                Case "02" ' CM Reg
                        ' (1) CM 開 / 關/停/ 換/變更速率
                        strCHcmdMsg = IIf(strOper_Type = "1", "CM 開機", "CM 關機")
                        Web_Command = mod_CMD_01.JustDoIt(strWork_ID, strCust_SO, strCustID, _
                                                                                            strCM_MAC, strEMTA_MAC, strClass_ID, _
                                                                                            strPC_IP_NO, strOper_Type)
                Case "12" ' ResetCM
                        ' (3) Reset CM
                        strCHcmdMsg = "Reset CM"
                        Web_Command = mod_CMD_03.JustDoIt(strWork_ID, strCust_SO, strCustID, strCM_MAC)
                        
                Case "20" ' PingCM
                        ' (4) Ping CM
                        strCHcmdMsg = "Ping CM"
                        Web_Command = mod_CMD_04.JustDoIt(strWork_ID, strCust_SO, strCustID, strCM_MAC)
                        
                Case "25" ' Clear CPE IP
                        ' (5) 釋放 動態IP
                        strCHcmdMsg = "釋放動態 IP"
                        Web_Command = mod_CMD_05.JustDoIt(strWork_ID, strCust_SO, strCustID, strCM_MAC)
                        
                Case "13" ' CMStatusQuery
                        ' (6) CM 狀態查詢
                        strCHcmdMsg = "查詢 CM 狀態"
                        Web_Command = mod_CMD_06.JustDoIt(strWork_ID, strCust_SO, strCustID, strCM_MAC)
                        
'                Case "24" ' CMCUSACCOUNT
'                        ' (7) ACS帳號密碼相關功能

                Case "21" ' HFCTypeQuery
                        ' (8) 查詢 HFC 服務類別
                        strCHcmdMsg = "查詢 HFC 服務類別"
                        If Get_Addr_Data(objCN, strComp_Code, lngAddrNo, varPara) Then
                            Web_Command = mod_CMD_08.JustDoIt(strWork_ID, strCust_SO, _
                                                                                                varPara(0, 0), varPara(1, 0) & "", _
                                                                                                varPara(2, 0) & "", varPara(3, 0) & "", _
                                                                                                varPara(4, 0) & "", varPara(5, 0) & "", _
                                                                                                varPara(6, 0) & "", varPara(7, 0) & "")
                        Else
                            Web_Command = False
                        End If
                Case "22" ' CMConnLog
                        ' (9) CM 連線紀錄查詢
                        strCHcmdMsg = "查詢 CM 連線紀錄"
                        Web_Command = mod_CMD_09.JustDoIt(strWork_ID, strCust_SO, strCustID, strCM_MAC, _
                                                                                            strStart_Date_Time, strEnd_Date_Time)
                        
                Case "23" ' CMQALog
                        ' (10) CM 連線品質查詢
                        strCHcmdMsg = "查詢 CM 連線品質"
                        Web_Command = mod_CMD_10.JustDoIt(strWork_ID, strCust_SO, strCustID, strCM_MAC, _
                                                                                            strStart_Date_Time, strEnd_Date_Time)
                        
                Case "26" ' CMPCLog
                        ' (11)PCIP連線記錄查詢
                        strCHcmdMsg = "查詢 PCIP 連線記錄"
                        Web_Command = mod_CMD_11.JustDoIt(strWork_ID, strCust_SO, strCustID, strPC_IP_NO, _
                                                                                            strStart_Date_Time, strEnd_Date_Time)
                        
                Case "27" ' CPIPReg
                        ' (2) CP IP 申請及退租
                        strCHcmdMsg = IIf(strOper_Type = "1", "CP IP 申請", "CP IP 退租")
                        If Get_Addr_Data(objCN, strComp_Code, lngAddrNo, varPara) Then
                            Web_Command = mod_CMD_02.JustDoIt(strWork_ID, strCust_SO, strCustID, strCM_MAC, _
                                                                                                strEMTA_MAC, strEMTA_Port, strCP_No, _
                                                                                                varPara(0, 0), varPara(1, 0) & "", _
                                                                                                varPara(2, 0) & "", varPara(3, 0) & "", _
                                                                                                varPara(4, 0) & "", varPara(5, 0) & "", _
                                                                                                varPara(6, 0) & "", varPara(7, 0) & "", _
                                                                                                strOper_Type)
                        Else
                            Web_Command = False
                        End If
                
                Case "24" ' Account
                        ' 帳號啟用停用
                
                Case "28" ' AccountSync
                        ' 新安裝帳號資料同步
    End Select
    
    Set objResult = Nothing
    
    ' Log 至 SO312 [ CM 帳號資料交換檔 ]
    If strID_Code = "24" Then
        ' Log 至 SO311 [ CMCP Gateway記錄檔 ]
        If LCase(strOper_Type) = "disable" Then
            strCHcmdMsg = "帳號停用"
        Else
            strCHcmdMsg = "帳號啟用"
        End If
        Insert_2_SO311_Para objCN, strID_Code, strWork_ID, strCust_SO, strCustID, strCM_MAC, _
                                            strEMTA_MAC, strEMTA_Port, strCP_No, strClass_ID, strPC_IP_NO, _
                                            strPC_IP, strOper_Type, strStart_Date_Time, strEnd_Date_Time, "", blnBatch
                                        
        Web_Command = Insert_2_SO312(objCN, strCust_SO, strCustID, strCM_MAC)
        If Not objPGB Is Nothing Then objPGB.Value = objPGB.Max
    
    Else
        If Web_Command Then
    '        ------------ 舊程式段 ------------
    '        Web_Command = mod_XMLhttp.Call_Web_Service("GET") ' 呼叫 Web Service 進行介接
    '        If Web_Command Then
    '        ------------ 舊程式段 ------------
    
            
            ' Log 至 SO311 [ CMCP Gateway記錄檔 ]
            Insert_2_SO311_Para objCN, strID_Code, strWork_ID, strCust_SO, strCustID, strCM_MAC, _
                                                strEMTA_MAC, strEMTA_Port, strCP_No, strClass_ID, strPC_IP_NO, _
                                                strPC_IP, strOper_Type, strStart_Date_Time, strEnd_Date_Time, _
                                                strPrcDate, blnBatch
                                                            
            If strPrcDate <> Empty Then
                Web_Command = True
                GoTo 66
            End If
            
            Dim strXMLdata As String
            Web_Command = Pooling(objCN, strCust_SO, strCustID, strXMLdata, objPGB)
            If Web_Command Then
                If mod_XML.ParseXML(strXMLdata) Then
                    With rsVirtual
                            If .RecordCount > 0 Then
                                .MoveFirst
                                Select Case strID_Code
                                            Case "13", "22", "23", "26"
                                                    strResult = .Fields("QueryResult") & ""
                                            Case Else
                                                    strResult = .Fields("OperResult") & ""
                                End Select
                                
                                strFaultReason = .Fields("FaultReason") & ""
                                
                                Upd_SO311_XML_Result objCN, strCust_SO, strCustID
                                
                                Set objResult = rsVirtual
                                If strID_Code = "13" Then ' (6) CM 狀態查詢
                                    ' 新增至 SO164 [ CM連線狀態 ]
                                    Web_Command = Insert_2_SO164(objCN, strCust_SO, strCustID, _
                                                                                            strCM_MAC, strWork_ID, lngAddrNo)
                                End If
                                
                            Else
                                strResult = "2"
                                strFaultReason = "無回應資料! (CS訊息)"
                            End If
                    End With
                End If
            Else
                strResult = ""
                strErrorMessage = "作業逾時 ( " & lngCMtimeout & " 秒 ) !!"
            End If
    '        ------------ 舊程式段 ------------
    '        Else
    '            strResult = "2"
    '            strFaultReason = "呼叫 Web Service 發生錯誤 !  (CS訊息)" ' & strErrorMessage
    ''            If Len(strFaultReason) > 30 Then strFaultReason = strErrorMessage
    '        End If
    '        ------------ 舊程式段 ------------
        Else
            strResult = "2"
            strFaultReason = "命令參數不足! 未呼叫 Web Service ..  (CS訊息)"
        End If
    End If
    
'sync: 裝機完工資料交換
'disable: 帳號停權
'Enable: 帳號啟用

'    Set objResult = CreateObject("ADODB.Recordset")
'    With objResult
'        .CursorLocation = 3
''        .Fields.Append "OperResult", adVarWChar, 2 ' 執行結果
''        .Fields.Append "FaultReason", adVarWChar, 60 ' 失敗原因
'        .Open
''        .Fields("FileDate").Properties("Optimize") = True
''        .Save stm, adPersistADTG
'    End With

' S_CmdSeqNo

66:
    On Error Resume Next
    objCN.Close
    Set objCN = Nothing

  Exit Function
ChkErr:
    ErrHandle "clsEMCcommand", "Web_Command"
End Function

Private Function Pooling(ByRef objCN As Object, _
                                        ByVal strCust_SO As String, _
                                        ByVal strCustID As String, _
                                        ByRef strXMLdata As String, _
                                        Optional ByRef objPGB As Object = Nothing) As Boolean
  On Error GoTo ChkErr
    
    Dim lngTimer As Long
    Dim intTickCount As Integer
    
    intTickCount = 0
    strXMLdata = Empty
    
    If Not objPGB Is Nothing Then objPGB.Max = lngCMtimeout
    Do Until intTickCount = lngCMtimeout
        If lngTimer = 0 Then lngTimer = GetTickCount
        If GetTickCount - lngTimer > 1000 Then
            lngTimer = GetTickCount
            If Not objPGB Is Nothing Then objPGB.Value = intTickCount
            strXMLdata = Get_XML_data(objCN, strCust_SO, strCustID)
            If strXMLdata <> Empty Then Exit Do
'            Debug.Print strXMLdata
            intTickCount = intTickCount + 1
            DoEvents
        End If
    Loop
    If Not objPGB Is Nothing Then objPGB.Value = lngCMtimeout
    
    Pooling = (Len(strXMLdata) > 0)
  
  Exit Function
ChkErr:
    ErrHandle "clsEMCcommand", "Pooling"
End Function

Private Function Get_XML_data(ByRef objCN As Object, _
                                                    ByVal strCust_SO As String, _
                                                    ByVal strCustID As String) As String
  On Error Resume Next
    Get_XML_data = objCN.Execute("SELECT XMLDATA FROM " & strCMowner & "SO311" & _
                                                        " WHERE " & IIf(strCustID = Empty, "", " CUSID='" & strCustID & "' AND ") & _
                                                        " CUSSO='" & strCust_SO & "' AND " & _
                                                        " CMDSEQNO='" & str_CmdSeqNo & "'").GetString(2, 1, "", "", "") & ""
    If Err.Number <> 0 Then Err.Clear
    Get_XML_data = RPxx(Get_XML_data)
End Function

Private Sub BGtrans(ByRef objCN As Object)
  On Error GoTo ChkErr
    objCN.BeginTrans
  Exit Sub
ChkErr:
    ErrHandle "clsEMCcommand", "BGtrans"
    objCN.RollbackTrans
End Sub

Private Sub CMtrans(ByRef objCN As Object)
  On Error GoTo ChkErr
    objCN.CommitTrans
  Exit Sub
ChkErr:
    ErrHandle "clsEMCcommand", "CMtrans"
    objCN.RollbackTrans
End Sub

' 新增 至 SO164 [ CM連線狀態 ]
Private Function Insert_2_SO164(ByRef objCN As Object, _
                                                        ByVal strCust_SO As String, _
                                                        Optional ByVal strCustID As String = "", _
                                                        Optional ByVal strCM_MAC As String = "", _
                                                        Optional ByVal strWork_ID As String = "", _
                                                        Optional ByVal lngAddrNo As Long) As Boolean
  On Error GoTo ChkErr
  
    Dim lngRcdAft As Long
    Dim cmd As Object
    
    Set cmd = CreateObject("ADODB.Command")
    BGtrans objCN ' Connection BeginTrans

    With cmd
            Set .ActiveConnection = objCN
            .CommandType = 1
            .CommandText = "INSERT INTO " & strCMowner & "SO164" & _
                                        " (COMPCODE,CUSTID,SNO,FACISNO,FACISEQNO,ENTRYTIME,RCMUPFREQ,RCMDOWNFREQ" & _
                                        ",UPLEVEL,DOWNLEVEL,RCMRECPW,RCMTRANSPW,RCMSNR,RCMTSRECPW,RCMTSRXSNR" & _
                                        ",RCMTSUPMOD,NODENO,CIRCUITNO,INSTADDRNO,UPDEN,UPDTIME)" & _
                                        " VALUES (?,?,?,?,?,SYSDATE,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
            '.Execute lngRcdAft, Array(strCust_SO, strCustID, "SNO", "FACISNO", "FACISEQNO", "ENTRYTIME"
            '"UPLEVEL", "DOWNLEVEL"
            .Execute lngRcdAft, Array(strCust_SO, strCustID, str_SNO, strCM_MAC, str_FaciSeqNo, _
                                                    GetNum(strR_CMUpFreq), GetNum(strR_CMDownFreg), Null, Null, _
                                                    GetNum(strR_CMRecPw), GetNum(strR_CMTransPw), _
                                                    GetNum(strR_CMSNR), GetNum(strR_CMTsRecPw), _
                                                    GetNum(strR_CMTsRxSNR), strR_CMTSUpMod, str_NodeNo, str_CircuitNo, _
                                                    lngAddrNo, strWork_ID, GetDTString(RightNow(objCN)))
            Insert_2_SO164 = (lngRcdAft = 1)
    End With
    
    CMtrans objCN ' Connection CommitTrans

    On Error Resume Next
    Set cmd = Nothing
  
  Exit Function
ChkErr:
    ErrHandle "clsEMCcommand", "Insert_2_SO164"
    objCN.RollbackTrans
End Function

Private Function GetNum(strValue As String) As String
  On Error GoTo ChkErr
    Dim intLoop As Integer, strTmp As String
    If Len(strValue) = 0 Then
        GetNum = ""
    Else
        For intLoop = 1 To Len(strValue)
          strTmp = Mid(strValue, intLoop, 1)
          If IsNumeric(strTmp) Or strTmp = "." Then GetNum = GetNum & strTmp
        Next
    End If
  Exit Function
ChkErr:
    ErrHandle "clsEMCcommand", "GetNum"
End Function

'Private Function RpMHz(strMHz As String) As String
'  On Error GoTo ChkErr
'    RpMHz = Replace(strMHz, "MHz", "", 1, , vbTextCompare)
' Exit Function
'ChkErr:
'    ErrHandle "clsEMCcommand", "RpMHz"
'End Function

' 新增 至 SO312 [ CM 帳號資料交換檔 ]
Private Function Insert_2_SO312(ByRef objCN As Object, _
                                                        ByVal strCust_SO As String, _
                                                        Optional ByVal strCustID As String = "", _
                                                        Optional ByVal strCM_MAC As String = "") As Boolean
  On Error GoTo ChkErr
  
    Dim lngRcdAft As Long
    Dim cmd As Object
    
    Set cmd = CreateObject("ADODB.Command")
    
    BGtrans objCN ' Connection BeginTrans
   
    With cmd
            Set .ActiveConnection = objCN
            .CommandType = 1
            .CommandText = "INSERT INTO " & strCMowner & "SO312 " & _
                                        "(COMPCODE,CUSTID,CMMAC,DIALACCOUNT,CMDSEQNO,COMMANDTYPE,ID)" & _
                                        " VALUES (?,?,?,?,?,?,?)"
                                   
'            系統台別 Cusso
'            客戶編號 Cusid
'            CM MAC  CM MAC
'            功能指令(開關)  OperType
'            撥接帳號 DialAccount
'            身份證字號  ID
'            指令 log table 流水號   CmdSeqno
            
            .Execute lngRcdAft, Array(strCust_SO, strCustID, strCM_MAC, str_DialAccount, str_CmdSeqNo, str_CmdType, str_ID)
            Insert_2_SO312 = (lngRcdAft = 1)
    End With
    
    CMtrans objCN ' Connection CommitTrans
    
    On Error Resume Next
    Set cmd = Nothing
    
  Exit Function
ChkErr:
    ErrHandle "clsEMCcommand", "Insert_2_SO312"
    objCN.RollbackTrans
End Function

' Update Web Service 結果至 SO311 [ CMCP Gateway記錄檔 ]
Private Sub Upd_SO311_XML_Result(ByRef objCN As Object, _
                                                            ByVal strCust_SO As String, _
                                                            Optional ByVal strCustID As String = "")

  On Error GoTo ChkErr
    Dim lngRcdAft As Long
    Dim cmd As Object
    
    Set cmd = CreateObject("ADODB.Command")
    
    BGtrans objCN ' Connection BeginTrans
    
    Set_Output_Para ' 設定 Web Service 回傳相關變數宣告

    With cmd
            Set .ActiveConnection = objCN
            .CommandType = 1
            .CommandText = "UPDATE " & strCMowner & "SO311 " & _
                                        " SET OPERRESULT=?,QUERYRESULT=?,FAULTREASON=?,GROUPID=?,CMSTATUS=?" & _
                                        ",RCMIP=?,RCMUPFREQ=?,RCMDOWNFREQ=?,RCMRECPW=?,RCMTRANSPW=?,RCMSNR=?" & _
                                        ",RCMONLINETIMES=?,RCMINOCTETS=?,RCMOUTOCTETS=?,RCMINERRORS=?" & _
                                        ",RCMOUTERRORS=?,RPCIP=?,RCMTSRECPW=?,RCMTSRXSNR=?,RCMTSUPMOD=?" & _
                                        ",RHFCTYPE=?,RPCOFFLINETIME=?,RPCONLINETIME=?,ROFFLINETIME=?,RONLINETIME=?" & _
                                        ",RDETCTTIME=?,RCMMAC=?,RCTIP=?,RNODEID=?" & _
                                        " WHERE CUSID=? AND CUSSO=? AND CMDSEQNO=?"
                                        
'                                        " WHERE CUSID='" & strCustID & "'" & _
                                        " AND CUSSO='" & strCust_SO & "'" & _
                                        " AND CMDSEQNO='" & str_CmdSeqNo & "'"
                                        
            On Error GoTo ChkErr
             .Execute lngRcdAft, Array(strOperResult, strQueryResult, strFaultReason, strGroupID, strCMStatus, _
                                                    strR_CMIP, strR_CMUpFreq, strR_CMDownFreg, strR_CMRecPw, strR_CMTransPw, _
                                                    strR_CMSNR, strR_CMOnLineTimes, strR_CMInOctets, strR_CMOutOtets, _
                                                    strR_CMInErrors, strR_CMOutErrors, strR_PCIP, strR_CMTsRecPw, _
                                                    strR_CMTsRxSNR, strR_CMTSUpMod, strR_HFCType, _
                                                    strR_PCOffLineTime, strR_PCOnLineTime, strR_OffLineTime, strR_OnLineTime, _
                                                    strR_DetctTime, strR_CMMAC, strR_CTIP, strR_NODEID, _
                                                    strCustID, strCust_SO, str_CmdSeqNo)
    End With
    
    CMtrans objCN ' Connection CommitTrans
    
    On Error Resume Next
    Set cmd = Nothing
  Exit Sub
ChkErr:
    ErrHandle "clsEMCcommand", "Upd_SO311_XML_Result"
    objCN.RollbackTrans
End Sub

' Insert Web Service 參數至 SO311 [ CMCP Gateway記錄檔 ]
Private Sub Insert_2_SO311_Para(ByRef objCN As Object, _
                                                        ByVal strID_Code As String, _
                                                        ByVal strWork_ID As String, _
                                                        ByVal strCust_SO As String, _
                                                        Optional ByVal strCustID As String = "", _
                                                        Optional ByVal strCM_MAC As String = "", _
                                                        Optional ByVal strEMTA_MAC As String = "", _
                                                        Optional ByVal strEMTA_Port As String = "", _
                                                        Optional ByVal strCP_No As String = "", _
                                                        Optional ByVal strClass_ID As String = "", _
                                                        Optional ByVal strPC_IP_NO As String = "", _
                                                        Optional ByVal strPC_IP As String = "", _
                                                        Optional ByVal strOper_Type As String = "", _
                                                        Optional ByVal strStart_Date_Time As String = "", _
                                                        Optional ByVal strEnd_Date_Time As String = "", _
                                                        Optional ByVal strPrcDate As String = "", _
                                                        Optional ByVal blnBatch As Boolean = False)
                                                    
  On Error GoTo ChkErr
    Dim lngRcdAft As Long
    Dim cmd As Object
    Dim rs311 As New ADODB.Recordset
    
    Set cmd = CreateObject("ADODB.Command")
    
'    ------------ 舊程式段 ------------
'    Set_Output_Para ' 設定 Web Service 回傳相關變數宣告
'    ------------ 舊程式段 ------------
    
    BGtrans objCN ' Connection BeginTrans

'    With rs311
'        If .State > 0 Then .Close
'        .CursorLocation = adUseClient
'        .Open "SELECT MSGCODE,MSGNAME,WORKID,CUSSO,CUSID,CMMAC,EMTAMAC,EMTAPORT" & _
'                    ",CPNO,CLASSID,PCIPNO,PCIP,OPERTYPE,STARTDATETIME,ENDDATETIME" & _
'                    ",ZONE,LIN,SECTION,LANE,ALLEY,SUBALLEY,NO1,NO2" & _
'                    ",EXECENTRY,EXECDATE,HOSTNAME,CMDSEQNO,DIALACCOUNT,DIALPASSWORD" & _
'                    ",CUSTNAME,FINTIME,TEL1,TEL2,COMMANDTYPE,ID,PROCESSINGDATE " & _
'                    " FROM " & strCMowner & "SO311 WHERE 0=1", objCN, adOpenKeyset, adLockOptimistic
'    End With
    
    With cmd
            Set .ActiveConnection = objCN
            .CommandType = 1
            .CommandText = "INSERT INTO " & strCMowner & "SO311 " & _
                                        "(MSGCODE,MSGNAME,WORKID,CUSSO,CUSID,CMMAC,EMTAMAC,EMTAPORT" & _
                                        ",CPNO,CLASSID,PCIPNO,PCIP,OPERTYPE,STARTDATETIME,ENDDATETIME" & _
                                        ",ZONE,LIN,SECTION,LANE,ALLEY,SUBALLEY,NO1,NO2" & _
                                        ",EXECENTRY,EXECDATE,HOSTNAME,CMDSEQNO,DIALACCOUNT,DIALPASSWORD" & _
                                        ",CUSTNAME,FINTIME,TEL1,TEL2,COMMANDTYPE,ID,PROCESSINGDATE)" & _
                                        " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?," & _
                                        YMDFmt & ",?,?,?,?," & YMDHMSFmt & ")"

'        ------------ 舊程式段 ------------
'            .CommandText = "INSERT INTO " & strOwner & "SO311 " & _
'                                        "(MSGCODE,MSGNAME,WORKID,CUSSO,CUSID,CMMAC,EMTAMAC,EMTAPORT" & _
'                                        ",CPNO,CLASSID,PCIPNO,PCIP,OPERTYPE,STARTDATETIME,ENDDATETIME" & _
'                                        ",ZONE,LIN,SECTION,LANE,ALLEY,SUBALLEY,NO1,NO2" & _
'                                        ",OPERRESULT,QUERYRESULT,FAULTREASON,GROUPID" & _
'                                        ",CMSTATUS,RCMIP,RCMUPFREQ,RCMDOWNFREQ,RCMRECPW,RCMTRANSPW,RCMSNR" & _
'                                        ",RCMONLINETIMES,RCMINOCTETS,RCMOUTOCTETS,RCMINERRORS,RCMOUTERRORS" & _
'                                        ",RPCIP,RCMTSRECPW,RCMTSRXSNR,RCMTSUPMOD,RHFCTYPE,RPCOFFLINETIME" & _
'                                        ",RPCONLINETIME,ROFFLINETIME,RONLINETIME,RDETCTTIME,RCMMAC" & _
'                                        ",RCTIP,RNODEID,EXECENTRY,EXECDATE,HOSTNAME" & _
'                                        ",CMDSEQNO,DIALACCOUNT,DIALPASSWORD,CUSTNAME,FINTIME" & _
'                                        ",TEL1,TEL2,COMMANDTYPE,ID)" & _
'                                        " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?," & _
'                                        "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
'        ------------ 舊程式段 ------------
                                   
            On Error Resume Next
            strCMD = Replace(strCMD, "/", "", 1)
            
            If InStr(1, strWork_ID, "=") Then
                strWork_ID = Split(strWork_ID, "=")(1)
                strCust_SO = Split(strCust_SO, "=")(1)
                strCustID = Split(strCustID, "=")(1)
                strCM_MAC = Split(strCM_MAC, "=")(1)
                strEMTA_MAC = Split(strEMTA_MAC, "=")(1)
                strEMTA_Port = Split(strEMTA_Port, "=")(1)
                strCP_No = Split(strCP_No, "=")(1)
                strClass_ID = Split(strClass_ID, "=")(1)
                strPC_IP_NO = Split(strPC_IP_NO, "=")(1)
                strPC_IP = Split(strPC_IP, "=")(1)
                strOper_Type = Split(strOper_Type, "=")(1)
                strStart_Date_Time = Split(strStart_Date_Time, "=")(1)
                strEnd_Date_Time = Split(strEnd_Date_Time, "=")(1)
            End If
            
            On Error GoTo ChkErr
            
            ' CmdSeqno          指令序號    VARCHAR2(20)    CmdSeqno指令序號
            ' (由FTP 收到結果檔後依此更新記錄TABLE的指令處理結果QueryResult)
            '                            00Yyyymmdd00000000
            '                            公司別+年月日+8碼流水序號
            ' DialAccount         撥接帳號    VARCHAR2(20)
            ' DialPassword      撥接密碼    VARCHAR2(24)
            
            
            '15. 由SO19F0A所下之整批指令
            '(1) 其SO311. ProcessingDate須依畫面條件上之預約時間來塞值。
            '(2) SO311.WorkID填為’Batch’ (表批次處理之指令，並與SO311.ExecEntry做區分)

              If strCHcmdMsg = Empty Then strCHcmdMsg = strCMD
              
'              rs311.AddNew _
'                    Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, _
'                            23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35), _
'                    Array(strID_Code, strCHcmdMsg, _
'                            IIf(blnBatch, "Batch", strWork_ID), _
'                            strCust_SO, _
'                            IIf(strCustID = "", Null, strCustID), _
'                            strCM_MAC, strEMTA_MAC, _
'                            IIf(strEMTA_Port = Empty, Null, strEMTA_Port), _
'                            strCP_No, strClass_ID, strPC_IP_NO, _
'                            strPC_IP, strOper_Type, strStart_Date_Time, strEnd_Date_Time, _
'                            strZone, strLin, strSection, strLane, strAlley, strSubAlley, strNo1, strNo2, _
'                            strWork_ID, GetDTString(RightNow(objCN)), Get_Computer_Name, _
'                            str_CmdSeqNo, str_DialAccount, str_DialPassword, str_CustName, _
'                            IIf(str_FinTime = Empty, Null, Format(str_FinTime, "YYYY/MM/DD")), _
'                            str_Tel1, str_Tel2, str_CmdType, str_ID, _
'                            IIf(strPrcDate = Empty, Null, Format(strPrcDate, "YYYY/MM/DD")))
              
             .Execute lngRcdAft, Array(strID_Code, strCHcmdMsg, _
                                                IIf(blnBatch, "Batch", strWork_ID), _
                                                strCust_SO, _
                                                IIf(strCustID = "", Null, strCustID), _
                                                strCM_MAC, strEMTA_MAC, _
                                                IIf(strEMTA_Port = Empty, Null, strEMTA_Port), _
                                                strCP_No, strClass_ID, strPC_IP_NO, _
                                                strPC_IP, strOper_Type, strStart_Date_Time, strEnd_Date_Time, _
                                                strZone, strLin, strSection, strLane, strAlley, strSubAlley, strNo1, strNo2, _
                                                strWork_ID, GetDTString(RightNow(objCN)), Get_Computer_Name, _
                                                str_CmdSeqNo, str_DialAccount, str_DialPassword, str_CustName, _
                                                IIf(str_FinTime = Empty, Null, Format(str_FinTime, "YYYY/MM/DD")), _
                                                str_Tel1, str_Tel2, str_CmdType, str_ID, _
                                                Format(strPrcDate, "YYYY/MM/DD HH:MM:SS"))
                                                
'        ------------ 舊程式段 ------------
'             .Execute lngRcdAft, Array(strID_Code, strCMD, strWork_ID, strCust_SO, _
                                                IIf(strCustID = "", Null, strCustID), _
                                                strCM_MAC, strEMTA_MAC, _
                                                IIf(strEMTA_Port = Empty, Null, strEMTA_Port), _
                                                strCP_No, strClass_ID, strPC_IP_NO, _
                                                strPC_IP, strOper_Type, strStart_Date_Time, strEnd_Date_Time, _
                                                strZone, strLin, strSection, strLane, strAlley, strSubAlley, strNo1, strNo2, _
                                                strOperResult, strQueryResult, strFaultReason, strGroupID, strCMStatus, _
                                                strR_CMIP, strR_CMUpFreq, strR_CMDownFreg, strR_CMRecPw, strR_CMTransPw, _
                                                strR_CMSNR, strR_CMOnLineTimes, strR_CMInOctets, strR_CMOutOtets, _
                                                strR_CMInErrors, strR_CMOutErrors, strR_PCIP, strR_CMTsRecPw, _
                                                strR_CMTsRxSNR, strR_CMTSUpMod, strR_HFCType, _
                                                strR_PCOffLineTime, strR_PCOnLineTime, strR_OffLineTime, strR_OnLineTime, _
                                                strR_DetctTime, strR_CMMAC, strR_CTIP, strR_NODEID, _
                                                strWork_ID, GetDTString(RightNow(objCN)), Get_Computer_Name, _
                                                str_CmdSeqNo, str_DialAccount, str_DialPassword, _
                                                str_CustName, IIf(str_FinTime = Empty, Null, str_FinTime), _
                                                str_Tel1, str_Tel2, str_CmdType, str_ID)
'        ------------ 舊程式段 ------------
    End With
    
    CMtrans objCN ' Connection CommitTrans
    
    On Error Resume Next
    Set cmd = Nothing
  
  Exit Sub
ChkErr:
    ErrHandle "clsEMCcommand", "Insert_2_SO311_Para"
    objCN.RollbackTrans
End Sub

' 取得地址相關資料以供傳送至 Web Service 進行參數傳遞
Private Function Get_Addr_Data(ByRef objCN As Object, _
                                                    ByRef strComp_Code As String, _
                                                    ByRef lngAddrNo As Long, _
                                                    ByRef varData As Variant) As Boolean
  On Error Resume Next
    varData = objCN.Execute("SELECT STRTNAME,LIN,SECTION,LANE,ALLEY,ALLEY2,NO1,NO2" & _
                                                " FROM " & strOwner & "SO014" & _
                                                " WHERE COMPCODE=" & strComp_Code & _
                                                " AND ADDRNO=" & lngAddrNo).GetRows
    Get_Addr_Data = (Err.Number = 0)
    If Not Get_Addr_Data Then
        Err.Clear
        strErrorMessage = "地址資料檔 [SO014] 查無 [地址編號] : " & lngAddrNo & " 的資料 !"
    Else
        Set_Addr_Data_Var varData ' 設定地址資料相關變數
    End If
End Function

' 設定地址資料相關變數
Private Sub Set_Addr_Data_Var(ByRef varData As Variant)
  On Error Resume Next
    strZone = varData(0, 0) & ""
    strLin = varData(1, 0) & ""
    strSection = varData(2, 0) & ""
    strLane = varData(3, 0) & ""
    strAlley = varData(4, 0) & ""
    strSubAlley = varData(5, 0) & ""
    strNo1 = varData(6, 0) & ""
    strNo2 = varData(7, 0) & ""
End Sub

' 清空地址資料相關變數
Private Sub Clear_Addr_Data_Var()
  On Error Resume Next
    strZone = Empty
    strLin = Empty
    strSection = Empty
    strLane = Empty
    strAlley = Empty
    strSubAlley = Empty
    strNo1 = Empty
    strNo2 = Empty
End Sub

' 設定 Web Service 回傳相關變數宣告
Private Sub Set_Output_Para()
  On Error Resume Next
    strOperResult = rsVirtual("OperResult") & ""
    strQueryResult = rsVirtual("QueryResult") & ""
    strFaultReason = rsVirtual("FaultReason") & ""
    strGroupID = rsVirtual("GROUPID") & ""
    strCMStatus = rsVirtual("CMStatus") & ""
    strR_CMIP = rsVirtual("CMIP") & ""
    strR_CMUpFreq = rsVirtual("CMUPFREQ") & ""
    strR_CMDownFreg = rsVirtual("CMDOWNFREQ") & ""
    strR_CMRecPw = rsVirtual("CMRECPW") & ""
    strR_CMTransPw = rsVirtual("CMTRANSPW") & ""
    strR_CMSNR = rsVirtual("CMSNR") & ""
    strR_CMOnLineTimes = rsVirtual("CMOnlineTimes") & ""
    strR_CMInOctets = rsVirtual("CMInOctets") & ""
    strR_CMOutOtets = rsVirtual("CMOutOctets") & ""
    strR_CMInErrors = rsVirtual("CMInErrors") & ""
    strR_CMOutErrors = rsVirtual("CMOutErrors") & ""
    strR_PCIP = rsVirtual("PCIP") & ""
    strR_CMTsRecPw = rsVirtual("CMTSRECPW") & ""
    strR_CMTsRxSNR = rsVirtual("CMTSRXSNR") & ""
    strR_CMTSUpMod = rsVirtual("CMTSUPMOD") & ""
    strR_HFCType = rsVirtual("HFCType")
    strR_PCOffLineTime = rsVirtual("PCOfflineTime") & ""
    strR_PCOnLineTime = rsVirtual("PCOnlineTime") & ""
    strR_OffLineTime = rsVirtual("OffLineTime") & ""
    strR_OnLineTime = rsVirtual("OnLineTime") & ""
    strR_DetctTime = rsVirtual("DetctTime") & ""
    strR_CMMAC = rsVirtual("CMMAC") & ""
    strR_CTIP = rsVirtual("CTIP") & ""
    strR_NODEID = rsVirtual("NODEID") & ""
End Sub
                                                        
' 清空 Web Service 回傳相關變數宣告
Private Sub Clear_Output_Para()
  On Error Resume Next
    strOperResult = Empty
    strQueryResult = Empty
    strFaultReason = Empty
    strGroupID = Empty
    strCMStatus = Empty
    strR_CMIP = Empty
    strR_CMUpFreq = Empty
    strR_CMDownFreg = Empty
    strR_CMRecPw = Empty
    strR_CMTransPw = Empty
    strR_CMSNR = Empty
    strR_CMOnLineTimes = Empty
    strR_CMInOctets = Empty
    strR_CMOutOtets = Empty
    strR_CMInErrors = Empty
    strR_CMOutErrors = Empty
    strR_PCIP = Empty
    strR_CMTsRecPw = Empty
    strR_CMTsRxSNR = Empty
    strR_CMTSUpMod = Empty
    strR_HFCType = Empty
    strR_PCOffLineTime = Empty
    strR_PCOnLineTime = Empty
    strR_OffLineTime = Empty
    strR_OnLineTime = Empty
    strR_DetctTime = Empty
    strR_CMMAC = Empty
    strR_CTIP = Empty
    strR_NODEID = Empty
End Sub

' 暫時提供測試 XML 用 , 以利測試
Public Function Return_XML_RS(Optional strXMLfile As String = "") As Object
  On Error GoTo ChkErr
    If strXMLfile <> Empty Then
        If mod_XML.ParseXML(strXMLfile) Then
            Set Return_XML_RS = rsVirtual
        Else
            Set Return_XML_RS = Nothing
        End If
    End If
  Exit Function
ChkErr:
    ErrHandle "clsEMCCommand", "Return_XML_RS"
End Function

' 取得錯誤訊息
Public Function GetErrMsg() As String
  On Error Resume Next
    GetErrMsg = strErrorMessage
End Function

Sub Class_Terminate()
  On Error Resume Next
  
    ' 地址相關資料變數宣告
    strZone = Empty
    strLin = Empty
    strSection = Empty
    strLane = Empty
    strAlley = Empty
    strSubAlley = Empty
    strNo1 = Empty
    strNo2 = Empty
    
    ' Web Service 回傳相關變數宣告
    strOperResult = Empty
    strQueryResult = Empty
    strFaultReason = Empty
    strGroupID = Empty
    strCMStatus = Empty
    strR_CMIP = Empty
    strR_CMUpFreq = Empty
    strR_CMDownFreg = Empty
    strR_CMRecPw = Empty
    strR_CMTransPw = Empty
    strR_CMSNR = Empty
    strR_CMOnLineTimes = Empty
    strR_CMInOctets = Empty
    strR_CMOutOtets = Empty
    strR_CMInErrors = Empty
    strR_CMOutErrors = Empty
    strR_PCIP = Empty
    strR_CMTsRecPw = Empty
    strR_CMTsRxSNR = Empty
    strR_CMTSUpMod = Empty
    strR_HFCType = Empty
    strR_PCOffLineTime = Empty
    strR_PCOnLineTime = Empty
    strR_OffLineTime = Empty
    strR_OnLineTime = Empty
    strR_DetctTime = Empty
    strR_CMMAC = Empty
    strR_CTIP = Empty
    strR_NODEID = Empty
    'strExecEntry =empty
    'strExecDate =empty
    'strHostName =empty
    
    str_DialAccount = Empty
    str_DialPassword = Empty
    str_Tel1 = Empty
    str_Tel2 = Empty
    str_CustName = Empty
    str_FinTime = Empty
    str_CmdSeqNo = Empty
    str_CmdType = Empty
    str_ID = Empty
    
    lngCMtimeout = Empty
    strCHcmdMsg = Empty
  
End Sub


'----------------------------------------------------------------------------------------
'|    回應結果                                           |                                                |
'----------------------------------------------------------------------------------------
'|    執行結果                                           |    OperResult / QueryResult    |
'----------------------------------------------------------------------------------------
'|    失敗原因                                           |    FaultReason                         |
'----------------------------------------------------------------------------------------
'|    GroupID                                          |    GroupID                               |
'----------------------------------------------------------------------------------------
'***|---靜態PCIP-----------------------------|----PCIPS--------------------------|*** 拿掉了
'----------------------------------------------------------------------------------------
'|    狀態                                                 |     CMStatus                            |
'----------------------------------------------------------------------------------------
'|    CMIP                                               |     CMIP                                   |
'----------------------------------------------------------------------------------------
'|    上行頻率                                           |    CMUpFreq                            |
'----------------------------------------------------------------------------------------
'|    下行頻率                                           |    CMDownFreq                       |
'----------------------------------------------------------------------------------------
'|    接收功率                                           |    CMRECPW                            |
'----------------------------------------------------------------------------------------
'|    發射功率                                           |    CMTransPW                         |
'----------------------------------------------------------------------------------------
'|    訊號雜訊比                                        |    CMSNR                                 |
'----------------------------------------------------------------------------------------
'|    開機時間                                           |    CMOnlineTimes                    |
'----------------------------------------------------------------------------------------
'|    上行累計封包                                     |    CMInOctets                         |
'----------------------------------------------------------------------------------------
'|    下行累計封包                                     |    CMOutOctets                       |
'----------------------------------------------------------------------------------------
'|    上行錯誤封包                                     |    CMInErrors                          |
'----------------------------------------------------------------------------------------
'|    下行錯誤封包                                     |    CMOutErrors                        |
'----------------------------------------------------------------------------------------
'|    PCIP                                                |    PCIP                                    |
'----------------------------------------------------------------------------------------
'|    RecPw                                             |    CMTSRECPW                        |
'----------------------------------------------------------------------------------------
'|    RxSNR                                             |    CMTSRXSNR                          |
'----------------------------------------------------------------------------------------
'|    調變方式                                           |    CMTSUPMOD                        |
'----------------------------------------------------------------------------------------
'|    服務種類                                           |    HFCType                              |
'----------------------------------------------------------------------------------------
'|    PC上線時間                                       |    PConlineTime                       |*** 新增
'----------------------------------------------------------------------------------------
'|    PC離線時間                                       |    PCofflineTime                       |*** 新增
'----------------------------------------------------------------------------------------
'|    上線時間                                           |    OnLineTime                          |
'----------------------------------------------------------------------------------------
'|    離線時間                                           |    OffLineTime                          |
'----------------------------------------------------------------------------------------
'|    偵測時間                                           |    DetctTime                             |
'----------------------------------------------------------------------------------------
'***|---客戶帳號------------------------------|----CUSACCO----------------------|*** 拿掉了
'----------------------------------------------------------------------------------------
'***|---客戶密碼------------------------------|----CUSPASSWORD---------------|*** 拿掉了
'----------------------------------------------------------------------------------------
'|    CMMAC                                           |    CMMAC                                 |
'----------------------------------------------------------------------------------------
'|    CTIP                                               |    CTIP                                     |
'----------------------------------------------------------------------------------------
'|    NODEID                                          |    NODEID                                |
'----------------------------------------------------------------------------------------
 
'<?xml version="1.0" encoding="utf-8"?>
'<definitions xmlns:http="http://schemas.xmlsoap.org/wsdl/http/" xmlns:soap="http://schemas.xmlsoap.org/wsdl/soap/" xmlns:s="http://www.w3.org/2001/XMLSchema" xmlns:s0="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW" xmlns:soapenc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:tm="http://microsoft.com/wsdl/mime/textMatching/" xmlns:mime="http://schemas.xmlsoap.org/wsdl/mime/" targetNamespace="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW" xmlns="http://schemas.xmlsoap.org/wsdl/">
'  <types>
'    <s:schema elementFormDefault="qualified" targetNamespace="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW">
'      <s:import namespace="http://www.w3.org/2001/XMLSchema" />
'      <s:element name="CMStatusQuery">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="WORKID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSSO" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CMMAC" type="s:string" />
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="CMStatusQueryResponse">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="CMStatusQueryResult">
'              <s:complexType>
'                <s:sequence>
'                  <s:element ref="s:schema" />
'                  <s:any />
'                </s:sequence>
'              </s:complexType>
'            </s:element>
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="ResetCM">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="WORKID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSSO" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CMMAC" type="s:string" />
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="ResetCMResponse">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="ResetCMResult">
'              <s:complexType>
'                <s:sequence>
'                  <s:element ref="s:schema" />
'                  <s:any />
'                </s:sequence>
'              </s:complexType>
'            </s:element>
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="ClearCPE">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="WORKID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSSO" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CMMAC" type="s:string" />
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="ClearCPEResponse">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="ClearCPEResult">
'              <s:complexType>
'                <s:sequence>
'                  <s:element ref="s:schema" />
'                  <s:any />
'                </s:sequence>
'              </s:complexType>
'            </s:element>
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="PingCM">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="WORKID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSSO" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CMMAC" type="s:string" />
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="PingCMResponse">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="PingCMResult">
'              <s:complexType>
'                <s:sequence>
'                  <s:element ref="s:schema" />
'                  <s:any />
'                </s:sequence>
'              </s:complexType>
'            </s:element>
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="HFCTypeQuery">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="WORKID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSSO" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="Zone" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="Lin" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="Section" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="Lane" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="Alley" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="SubAlley" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="NO1" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="NO2" type="s:string" />
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="HFCTypeQueryResponse">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="HFCTypeQueryResult">
'              <s:complexType>
'                <s:sequence>
'                  <s:element ref="s:schema" />
'                  <s:any />
'                </s:sequence>
'              </s:complexType>
'            </s:element>
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="CPIPReg">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="WORKID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSSO" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CMMAC" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="EMTAMAC" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="EMTAPort" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CPNO" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="Zone" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="Lin" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="Section" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="Lane" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="Alley" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="SubAlley" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="NO1" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="NO2" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="OPERTYPE" type="s:string" />
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="CPIPRegResponse">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="CPIPRegResult">
'              <s:complexType>
'                <s:sequence>
'                  <s:element ref="s:schema" />
'                  <s:any />
'                </s:sequence>
'              </s:complexType>
'            </s:element>
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="CMReg">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="WORKID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSSO" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CMMAC" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="EMTAMAC" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CLASSID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="PCIPNO" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="OPERTYPE" type="s:string" />
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="CMRegResponse">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="CMRegResult">
'              <s:complexType>
'                <s:sequence>
'                  <s:element ref="s:schema" />
'                  <s:any />
'                </s:sequence>
'              </s:complexType>
'            </s:element>
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="CMConnLog">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="WORKID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSSO" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CMMAC" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="StartDateTime" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="EndDateTime" type="s:string" />
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="CMConnLogResponse">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="CMConnLogResult">
'              <s:complexType>
'                <s:sequence>
'                  <s:element ref="s:schema" />
'                  <s:any />
'                </s:sequence>
'              </s:complexType>
'            </s:element>
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="CMQALog">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="WORKID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSSO" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CMMAC" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="StartDateTime" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="EndDateTime" type="s:string" />
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="CMQALogResponse">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="CMQALogResult">
'              <s:complexType>
'                <s:sequence>
'                  <s:element ref="s:schema" />
'                  <s:any />
'                </s:sequence>
'              </s:complexType>
'            </s:element>
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="CMPCLog">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="WORKID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSSO" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="CUSID" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="PCIP" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="StartDateTime" type="s:string" />
'            <s:element minOccurs="0" maxOccurs="1" name="EndDateTime" type="s:string" />
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="CMPCLogResponse">
'        <s:complexType>
'          <s:sequence>
'            <s:element minOccurs="0" maxOccurs="1" name="CMPCLogResult">
'              <s:complexType>
'                <s:sequence>
'                  <s:element ref="s:schema" />
'                  <s:any />
'                </s:sequence>
'              </s:complexType>
'            </s:element>
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'      <s:element name="DataSet" nillable="true">
'        <s:complexType>
'          <s:sequence>
'            <s:element ref="s:schema" />
'            <s:any />
'          </s:sequence>
'        </s:complexType>
'      </s:element>
'    </s:schema>
'  </types>
'  <message name="CMStatusQuerySoapIn">
'    <part name="parameters" element="s0:CMStatusQuery" />
'  </message>
'  <message name="CMStatusQuerySoapOut">
'    <part name="parameters" element="s0:CMStatusQueryResponse" />
'  </message>
'  <message name="ResetCMSoapIn">
'    <part name="parameters" element="s0:ResetCM" />
'  </message>
'  <message name="ResetCMSoapOut">
'    <part name="parameters" element="s0:ResetCMResponse" />
'  </message>
'  <message name="ClearCPESoapIn">
'    <part name="parameters" element="s0:ClearCPE" />
'  </message>
'  <message name="ClearCPESoapOut">
'    <part name="parameters" element="s0:ClearCPEResponse" />
'  </message>
'  <message name="PingCMSoapIn">
'    <part name="parameters" element="s0:PingCM" />
'  </message>
'  <message name="PingCMSoapOut">
'    <part name="parameters" element="s0:PingCMResponse" />
'  </message>
'  <message name="HFCTypeQuerySoapIn">
'    <part name="parameters" element="s0:HFCTypeQuery" />
'  </message>
'  <message name="HFCTypeQuerySoapOut">
'    <part name="parameters" element="s0:HFCTypeQueryResponse" />
'  </message>
'  <message name="CPIPRegSoapIn">
'    <part name="parameters" element="s0:CPIPReg" />
'  </message>
'  <message name="CPIPRegSoapOut">
'    <part name="parameters" element="s0:CPIPRegResponse" />
'  </message>
'  <message name="CMRegSoapIn">
'    <part name="parameters" element="s0:CMReg" />
'  </message>
'  <message name="CMRegSoapOut">
'    <part name="parameters" element="s0:CMRegResponse" />
'  </message>
'  <message name="CMConnLogSoapIn">
'    <part name="parameters" element="s0:CMConnLog" />
'  </message>
'  <message name="CMConnLogSoapOut">
'    <part name="parameters" element="s0:CMConnLogResponse" />
'  </message>
'  <message name="CMQALogSoapIn">
'    <part name="parameters" element="s0:CMQALog" />
'  </message>
'  <message name="CMQALogSoapOut">
'    <part name="parameters" element="s0:CMQALogResponse" />
'  </message>
'  <message name="CMPCLogSoapIn">
'    <part name="parameters" element="s0:CMPCLog" />
'  </message>
'  <message name="CMPCLogSoapOut">
'    <part name="parameters" element="s0:CMPCLogResponse" />
'  </message>
'  <message name="CMStatusQueryHttpGetIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'  </message>
'  <message name="CMStatusQueryHttpGetOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="ResetCMHttpGetIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'  </message>
'  <message name="ResetCMHttpGetOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="ClearCPEHttpGetIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'  </message>
'  <message name="ClearCPEHttpGetOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="PingCMHttpGetIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'  </message>
'  <message name="PingCMHttpGetOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="HFCTypeQueryHttpGetIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="Zone" type="s:string" />
'    <part name="Lin" type="s:string" />
'    <part name="Section" type="s:string" />
'    <part name="Lane" type="s:string" />
'    <part name="Alley" type="s:string" />
'    <part name="SubAlley" type="s:string" />
'    <part name="NO1" type="s:string" />
'    <part name="NO2" type="s:string" />
'  </message>
'  <message name="HFCTypeQueryHttpGetOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="CPIPRegHttpGetIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'    <part name="EMTAMAC" type="s:string" />
'    <part name="EMTAPort" type="s:string" />
'    <part name="CPNO" type="s:string" />
'    <part name="Zone" type="s:string" />
'    <part name="Lin" type="s:string" />
'    <part name="Section" type="s:string" />
'    <part name="Lane" type="s:string" />
'    <part name="Alley" type="s:string" />
'    <part name="SubAlley" type="s:string" />
'    <part name="NO1" type="s:string" />
'    <part name="NO2" type="s:string" />
'    <part name="OPERTYPE" type="s:string" />
'  </message>
'  <message name="CPIPRegHttpGetOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="CMRegHttpGetIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'    <part name="EMTAMAC" type="s:string" />
'    <part name="CLASSID" type="s:string" />
'    <part name="PCIPNO" type="s:string" />
'    <part name="OPERTYPE" type="s:string" />
'  </message>
'  <message name="CMRegHttpGetOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="CMConnLogHttpGetIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'    <part name="StartDateTime" type="s:string" />
'    <part name="EndDateTime" type="s:string" />
'  </message>
'  <message name="CMConnLogHttpGetOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="CMQALogHttpGetIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'    <part name="StartDateTime" type="s:string" />
'    <part name="EndDateTime" type="s:string" />
'  </message>
'  <message name="CMQALogHttpGetOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="CMPCLogHttpGetIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="PCIP" type="s:string" />
'    <part name="StartDateTime" type="s:string" />
'    <part name="EndDateTime" type="s:string" />
'  </message>
'  <message name="CMPCLogHttpGetOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="CMStatusQueryHttpPostIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'  </message>
'  <message name="CMStatusQueryHttpPostOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="ResetCMHttpPostIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'  </message>
'  <message name="ResetCMHttpPostOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="ClearCPEHttpPostIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'  </message>
'  <message name="ClearCPEHttpPostOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="PingCMHttpPostIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'  </message>
'  <message name="PingCMHttpPostOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="HFCTypeQueryHttpPostIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="Zone" type="s:string" />
'    <part name="Lin" type="s:string" />
'    <part name="Section" type="s:string" />
'    <part name="Lane" type="s:string" />
'    <part name="Alley" type="s:string" />
'    <part name="SubAlley" type="s:string" />
'    <part name="NO1" type="s:string" />
'    <part name="NO2" type="s:string" />
'  </message>
'  <message name="HFCTypeQueryHttpPostOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="CPIPRegHttpPostIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'    <part name="EMTAMAC" type="s:string" />
'    <part name="EMTAPort" type="s:string" />
'    <part name="CPNO" type="s:string" />
'    <part name="Zone" type="s:string" />
'    <part name="Lin" type="s:string" />
'    <part name="Section" type="s:string" />
'    <part name="Lane" type="s:string" />
'    <part name="Alley" type="s:string" />
'    <part name="SubAlley" type="s:string" />
'    <part name="NO1" type="s:string" />
'    <part name="NO2" type="s:string" />
'    <part name="OPERTYPE" type="s:string" />
'  </message>
'  <message name="CPIPRegHttpPostOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="CMRegHttpPostIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'    <part name="EMTAMAC" type="s:string" />
'    <part name="CLASSID" type="s:string" />
'    <part name="PCIPNO" type="s:string" />
'    <part name="OPERTYPE" type="s:string" />
'  </message>
'  <message name="CMRegHttpPostOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="CMConnLogHttpPostIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'    <part name="StartDateTime" type="s:string" />
'    <part name="EndDateTime" type="s:string" />
'  </message>
'  <message name="CMConnLogHttpPostOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="CMQALogHttpPostIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="CMMAC" type="s:string" />
'    <part name="StartDateTime" type="s:string" />
'    <part name="EndDateTime" type="s:string" />
'  </message>
'  <message name="CMQALogHttpPostOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <message name="CMPCLogHttpPostIn">
'    <part name="WORKID" type="s:string" />
'    <part name="CUSSO" type="s:string" />
'    <part name="CUSID" type="s:string" />
'    <part name="PCIP" type="s:string" />
'    <part name="StartDateTime" type="s:string" />
'    <part name="EndDateTime" type="s:string" />
'  </message>
'  <message name="CMPCLogHttpPostOut">
'    <part name="Body" element="s0:DataSet" />
'  </message>
'  <portType name="EMCWebServiceGWSoap">
'    <operation name="CMStatusQuery">
'      <input message="s0:CMStatusQuerySoapIn" />
'      <output message="s0:CMStatusQuerySoapOut" />
'    </operation>
'    <operation name="ResetCM">
'      <input message="s0:ResetCMSoapIn" />
'      <output message="s0:ResetCMSoapOut" />
'    </operation>
'    <operation name="ClearCPE">
'      <input message="s0:ClearCPESoapIn" />
'      <output message="s0:ClearCPESoapOut" />
'    </operation>
'    <operation name="PingCM">
'      <input message="s0:PingCMSoapIn" />
'      <output message="s0:PingCMSoapOut" />
'    </operation>
'    <operation name="HFCTypeQuery">
'      <input message="s0:HFCTypeQuerySoapIn" />
'      <output message="s0:HFCTypeQuerySoapOut" />
'    </operation>
'    <operation name="CPIPReg">
'      <input message="s0:CPIPRegSoapIn" />
'      <output message="s0:CPIPRegSoapOut" />
'    </operation>
'    <operation name="CMReg">
'      <input message="s0:CMRegSoapIn" />
'      <output message="s0:CMRegSoapOut" />
'    </operation>
'    <operation name="CMConnLog">
'      <input message="s0:CMConnLogSoapIn" />
'      <output message="s0:CMConnLogSoapOut" />
'    </operation>
'    <operation name="CMQALog">
'      <input message="s0:CMQALogSoapIn" />
'      <output message="s0:CMQALogSoapOut" />
'    </operation>
'    <operation name="CMPCLog">
'      <input message="s0:CMPCLogSoapIn" />
'      <output message="s0:CMPCLogSoapOut" />
'    </operation>
'  </portType>
'  <portType name="EMCWebServiceGWHttpGet">
'    <operation name="CMStatusQuery">
'      <input message="s0:CMStatusQueryHttpGetIn" />
'      <output message="s0:CMStatusQueryHttpGetOut" />
'    </operation>
'    <operation name="ResetCM">
'      <input message="s0:ResetCMHttpGetIn" />
'      <output message="s0:ResetCMHttpGetOut" />
'    </operation>
'    <operation name="ClearCPE">
'      <input message="s0:ClearCPEHttpGetIn" />
'      <output message="s0:ClearCPEHttpGetOut" />
'    </operation>
'    <operation name="PingCM">
'      <input message="s0:PingCMHttpGetIn" />
'      <output message="s0:PingCMHttpGetOut" />
'    </operation>
'    <operation name="HFCTypeQuery">
'      <input message="s0:HFCTypeQueryHttpGetIn" />
'      <output message="s0:HFCTypeQueryHttpGetOut" />
'    </operation>
'    <operation name="CPIPReg">
'      <input message="s0:CPIPRegHttpGetIn" />
'      <output message="s0:CPIPRegHttpGetOut" />
'    </operation>
'    <operation name="CMReg">
'      <input message="s0:CMRegHttpGetIn" />
'      <output message="s0:CMRegHttpGetOut" />
'    </operation>
'    <operation name="CMConnLog">
'      <input message="s0:CMConnLogHttpGetIn" />
'      <output message="s0:CMConnLogHttpGetOut" />
'    </operation>
'    <operation name="CMQALog">
'      <input message="s0:CMQALogHttpGetIn" />
'      <output message="s0:CMQALogHttpGetOut" />
'    </operation>
'    <operation name="CMPCLog">
'      <input message="s0:CMPCLogHttpGetIn" />
'      <output message="s0:CMPCLogHttpGetOut" />
'    </operation>
'  </portType>
'  <portType name="EMCWebServiceGWHttpPost">
'    <operation name="CMStatusQuery">
'      <input message="s0:CMStatusQueryHttpPostIn" />
'      <output message="s0:CMStatusQueryHttpPostOut" />
'    </operation>
'    <operation name="ResetCM">
'      <input message="s0:ResetCMHttpPostIn" />
'      <output message="s0:ResetCMHttpPostOut" />
'    </operation>
'    <operation name="ClearCPE">
'      <input message="s0:ClearCPEHttpPostIn" />
'      <output message="s0:ClearCPEHttpPostOut" />
'    </operation>
'    <operation name="PingCM">
'      <input message="s0:PingCMHttpPostIn" />
'      <output message="s0:PingCMHttpPostOut" />
'    </operation>
'    <operation name="HFCTypeQuery">
'      <input message="s0:HFCTypeQueryHttpPostIn" />
'      <output message="s0:HFCTypeQueryHttpPostOut" />
'    </operation>
'    <operation name="CPIPReg">
'      <input message="s0:CPIPRegHttpPostIn" />
'      <output message="s0:CPIPRegHttpPostOut" />
'    </operation>
'    <operation name="CMReg">
'      <input message="s0:CMRegHttpPostIn" />
'      <output message="s0:CMRegHttpPostOut" />
'    </operation>
'    <operation name="CMConnLog">
'      <input message="s0:CMConnLogHttpPostIn" />
'      <output message="s0:CMConnLogHttpPostOut" />
'    </operation>
'    <operation name="CMQALog">
'      <input message="s0:CMQALogHttpPostIn" />
'      <output message="s0:CMQALogHttpPostOut" />
'    </operation>
'    <operation name="CMPCLog">
'      <input message="s0:CMPCLogHttpPostIn" />
'      <output message="s0:CMPCLogHttpPostOut" />
'    </operation>
'  </portType>
'  <binding name="EMCWebServiceGWSoap" type="s0:EMCWebServiceGWSoap">
'    <soap:binding transport="http://schemas.xmlsoap.org/soap/http" style="document" />
'    <operation name="CMStatusQuery">
'      <soap:operation soapAction="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/CMStatusQuery" style="document" />
'      <input>
'        <soap:body use="literal" />
'      </input>
'      <output>
'        <soap:body use="literal" />
'      </output>
'    </operation>
'    <operation name="ResetCM">
'      <soap:operation soapAction="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/ResetCM" style="document" />
'      <input>
'        <soap:body use="literal" />
'      </input>
'      <output>
'        <soap:body use="literal" />
'      </output>
'    </operation>
'    <operation name="ClearCPE">
'      <soap:operation soapAction="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/ClearCPE" style="document" />
'      <input>
'        <soap:body use="literal" />
'      </input>
'      <output>
'        <soap:body use="literal" />
'      </output>
'    </operation>
'    <operation name="PingCM">
'      <soap:operation soapAction="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/PingCM" style="document" />
'      <input>
'        <soap:body use="literal" />
'      </input>
'      <output>
'        <soap:body use="literal" />
'      </output>
'    </operation>
'    <operation name="HFCTypeQuery">
'      <soap:operation soapAction="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/HFCTypeQuery" style="document" />
'      <input>
'        <soap:body use="literal" />
'      </input>
'      <output>
'        <soap:body use="literal" />
'      </output>
'    </operation>
'    <operation name="CPIPReg">
'      <soap:operation soapAction="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/CPIPReg" style="document" />
'      <input>
'        <soap:body use="literal" />
'      </input>
'      <output>
'        <soap:body use="literal" />
'      </output>
'    </operation>
'    <operation name="CMReg">
'      <soap:operation soapAction="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/CMReg" style="document" />
'      <input>
'        <soap:body use="literal" />
'      </input>
'      <output>
'        <soap:body use="literal" />
'      </output>
'    </operation>
'    <operation name="CMConnLog">
'      <soap:operation soapAction="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/CMConnLog" style="document" />
'      <input>
'        <soap:body use="literal" />
'      </input>
'      <output>
'        <soap:body use="literal" />
'      </output>
'    </operation>
'    <operation name="CMQALog">
'      <soap:operation soapAction="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/CMQALog" style="document" />
'      <input>
'        <soap:body use="literal" />
'      </input>
'      <output>
'        <soap:body use="literal" />
'      </output>
'    </operation>
'    <operation name="CMPCLog">
'      <soap:operation soapAction="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW/CMPCLog" style="document" />
'      <input>
'        <soap:body use="literal" />
'      </input>
'      <output>
'        <soap:body use="literal" />
'      </output>
'    </operation>
'  </binding>
'  <binding name="EMCWebServiceGWHttpGet" type="s0:EMCWebServiceGWHttpGet">
'    <http:binding verb="GET" />
'    <operation name="CMStatusQuery">
'      <http:operation location="/CMStatusQuery" />
'      <input>
'        <http:urlEncoded />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="ResetCM">
'      <http:operation location="/ResetCM" />
'      <input>
'        <http:urlEncoded />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="ClearCPE">
'      <http:operation location="/ClearCPE" />
'      <input>
'        <http:urlEncoded />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="PingCM">
'      <http:operation location="/PingCM" />
'      <input>
'        <http:urlEncoded />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="HFCTypeQuery">
'      <http:operation location="/HFCTypeQuery" />
'      <input>
'        <http:urlEncoded />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="CPIPReg">
'      <http:operation location="/CPIPReg" />
'      <input>
'        <http:urlEncoded />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="CMReg">
'      <http:operation location="/CMReg" />
'      <input>
'        <http:urlEncoded />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="CMConnLog">
'      <http:operation location="/CMConnLog" />
'      <input>
'        <http:urlEncoded />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="CMQALog">
'      <http:operation location="/CMQALog" />
'      <input>
'        <http:urlEncoded />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="CMPCLog">
'      <http:operation location="/CMPCLog" />
'      <input>
'        <http:urlEncoded />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'  </binding>
'  <binding name="EMCWebServiceGWHttpPost" type="s0:EMCWebServiceGWHttpPost">
'    <http:binding verb="POST" />
'    <operation name="CMStatusQuery">
'      <http:operation location="/CMStatusQuery" />
'      <input>
'        <mime:content type="application/x-www-form-urlencoded" />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="ResetCM">
'      <http:operation location="/ResetCM" />
'      <input>
'        <mime:content type="application/x-www-form-urlencoded" />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="ClearCPE">
'      <http:operation location="/ClearCPE" />
'      <input>
'        <mime:content type="application/x-www-form-urlencoded" />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="PingCM">
'      <http:operation location="/PingCM" />
'      <input>
'        <mime:content type="application/x-www-form-urlencoded" />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="HFCTypeQuery">
'      <http:operation location="/HFCTypeQuery" />
'      <input>
'        <mime:content type="application/x-www-form-urlencoded" />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="CPIPReg">
'      <http:operation location="/CPIPReg" />
'      <input>
'        <mime:content type="application/x-www-form-urlencoded" />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="CMReg">
'      <http:operation location="/CMReg" />
'      <input>
'        <mime:content type="application/x-www-form-urlencoded" />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="CMConnLog">
'      <http:operation location="/CMConnLog" />
'      <input>
'        <mime:content type="application/x-www-form-urlencoded" />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="CMQALog">
'      <http:operation location="/CMQALog" />
'      <input>
'        <mime:content type="application/x-www-form-urlencoded" />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'    <operation name="CMPCLog">
'      <http:operation location="/CMPCLog" />
'      <input>
'        <mime:content type="application/x-www-form-urlencoded" />
'      </input>
'      <output>
'        <mime:mimeXml part="Body" />
'      </output>
'    </operation>
'  </binding>
'  <service name="EMCWebServiceGW">
'    <port name="EMCWebServiceGWSoap" binding="s0:EMCWebServiceGWSoap">
'      <soap:address location="http://210.202.146.97:9090/CMWebServiceGWDemo/EMCWebServiceGW.asmx" />
'    </port>
'    <port name="EMCWebServiceGWHttpGet" binding="s0:EMCWebServiceGWHttpGet">
'      <http:address location="http://210.202.146.97:9090/CMWebServiceGWDemo/EMCWebServiceGW.asmx" />
'    </port>
'    <port name="EMCWebServiceGWHttpPost" binding="s0:EMCWebServiceGWHttpPost">
'      <http:address location="http://210.202.146.97:9090/CMWebServiceGWDemo/EMCWebServiceGW.asmx" />
'    </port>
'  </service>
'</definitions>
