VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommand"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private cn As Object
Private strQry As String

'CmdDisp = Switch(Value = "C1", "裝機 (CM)", Value = "E1", "裝機 (EMTA)", _
'                                Value = "C2", "拆機 (CM裝機退件)", Value = "E2", "拆機 (EMTA裝機退件)", _
'                                Value = "E3", "CP 裝機", Value = "A1_P", "CP Only 拆機", _
'                                Value = "E4", "EMTA 加裝 CP 門號", Value = "E5", "EMTA 拆 CP 門號", _
'                                Value = "A1_D", "軟拆", Value = "A1_C", "軟復", _
'                                Value = "A1_BU", "速率升降級", Value = "A1_PIP", "申請動態IP", _
'                                Value = "A1_MIP", "取消動態IP", Value = "Q1", "CM reset", _
'                                Value = "Q2", "查詢CM狀況", Value = "A7", "CM 拆舊 / 換新", _
'                                Value = "A8", "EMTA 拆舊 / 換新", Value = "A9", "拆舊CM / 換新EMTA", _
'                                Value = "A10", "拆舊EMTA / 換新 CM", Value = "A6", "變更 CPE MAC", _
'                                Value = "A5", "CM IP  CPE IP 設定", Value = "A4", "EMTA IP Priv IP 設定", _
'                                Value = "A3", "取消固定IP", Value = "A2", "申請固定IP", _
'                                Value = "C6", "CM 移機", Value = "E6", "EMTA 移機")

'If Not blnTransation Then gcnGi.BeginTrans
'Public strResvTime As String

Public Function ExecCmd(ByRef Gcn As Object, ByRef rs04 As Object, CmdID As String, _
                                        AryGi As String, ErrPath As String, Optional ResvTime As String = "", _
                                        Optional Transaction As Boolean = True, Optional RetMsg As String = "", _
                                        Optional FaciCode As String = "", Optional FaciName As String = "", _
                                        Optional FaciSno As String = "", Optional SNO As String = "", _
                                        Optional MediaBillNo As String = "", Optional CmdSeqNo As String = "", _
                                        Optional BaudRate As String = "", Optional ProbStopDate As String = "", _
                                        Optional ClearSNo As Boolean = False, _
                                        Optional ByRef rsNew As ADODB.Recordset = Nothing, _
                                        Optional ByVal strNewBaud As String = Empty, _
                                        Optional ByVal strNewDynaIPcnt As String = Empty, _
                                        Optional ByVal strNewFixIPcnt As String = Empty, _
                                        Optional ByVal strNewIPAddress As String = Empty, _
                                        Optional ByVal strNewModel As String = Empty, _
                                        Optional ByVal strNewFaci As String = Empty) As Boolean
  On Error GoTo ChkErr
  
    ExecCmd = False
    UI_Mode = False
    blnNoShowMsg = False
    RetMsg = ""
    strRetMsg = ""
    
    If Gcn Is Nothing Then RetMsg = "請傳入Connection 連線物件 !": Exit Function
    If Gcn.State <= 0 Then RetMsg = "請確認 Connection 連線物件已開啟 !": Exit Function
    If rs04 Is Nothing Then RetMsg = "請傳入 SO004 Recordset 資料錄集物件 !": Exit Function
    If rs04.State <= 0 Then RetMsg = "請確認 SO004 Recordset 資料錄集已開啟 !": Exit Function
    If rs04.EOF Then RetMsg = "請確認 SO004 Recordset 資料錄集中有資料 !": Exit Function
    If Len(AryGi) = 0 Then RetMsg = "請傳入 GaryGi 字串 !": Exit Function
    If Len(ErrPath) = 0 Then RetMsg = "請傳入 ErrPath 路徑 ! ( Error Log Path )": Exit Function
    
    Dim intComp As Integer
    Dim lngCustId As Long
    Dim SvcTp As String
    Dim rsTmp As New ADODB.Recordset
    
    'Set cn = Gcn
    Set gcnGi = Gcn
    GetGlobal AryGi
    gErrLogPath = ErrPath
    
    Set objCmdCN = CreateObject("ADODB.Connection")
    With objCmdCN
        .CursorLocation = 3
        .Open gcnGi.ConnectionString & ";Persist Security Info=True"
    End With
    '#6066 如果rsNew有傳值,則rsUPD=rsNew By Kin 2011/07/26
    If strNewBaud <> Empty Then strNewBaudRate = strNewBaud
    If strNewDynaIPcnt <> Empty Then strDynaIPcnt = strNewDynaIPcnt
    If strNewFixIPcnt <> Empty Then strFixIPcnt = strNewFixIPcnt
    If strNewIPAddress <> Empty Then strNewIP = strNewIPAddress
    If strNewModel <> Empty Then strNewModelCode = strNewModel
    If strNewFaci <> Empty Then strNewFaciSno = strNewFaci
    
'    Set rsUpd = Nothing
'    If Not rsNew Is Nothing Then
'        If rsNew.State = adStateOpen Then
'            If rsNew.RecordCount > 0 Then
'                Set rsUpd = rsNew
'            End If
'        End If
'    End If
    
    
    intComp = Val(rs04!CompCode & "")
    lngCustId = Val(rs04!CustID & "")
    SvcTp = rs04!ServiceType & ""
    
    If ResvTime <> "" Then
        If InStr(1, ResvTime, "/") <= 0 Then ResvTime = Format(ResvTime, "####/##/## ##:##:##")
    End If
    
    If ProbStopDate <> "" Then
        If InStr(1, ProbStopDate, "/") <= 0 Then ProbStopDate = Format(ProbStopDate, "####/##/##")
    End If
    
    strResvTime = ResvTime
    strProbStopDate = ProbStopDate
    
    CmdID = UCase(CmdID)
    
    If intComp = 0 Then RetMsg = "SO004 資料中 [公司別] 為 0 不正確 !": Exit Function
    If lngCustId = 0 Then RetMsg = "SO004 資料中 [客編] 為 0 不正確 !": Exit Function
    If Len(SvcTp) = 0 Then RetMsg = "SO004 資料中 [服務別] 為空 不正確 !": Exit Function

    gCompFilterStr = intComp
    gDefaultComp = intComp
    gCompCode = intComp
    gServiceType = SvcTp
    
    strSno = SNO
    strMediaBillNo = MediaBillNo
    strBaudRate = BaudRate
    
    int_EMTA_IP_Type = Val(GetSystemParaItem("EMTAIPTYPE", CStr(intComp), SvcTp, "SO041", , True) & "")
    
    strQry = "SELECT INSTADDRNO FROM " & GetOwner & "SO001" & _
                    " WHERE CUSTID = " & lngCustId & " AND COMPCODE = " & intComp
    
    If Not GetRS(rsTmp, strQry, gcnGi, , , , , , , True) Then
        RetMsg = "SO001 查詢 [公司別] 為 " & intComp & " , [客編] 為 " & lngCustId & " 的 [地址編號] 資料失敗 !"
        Exit Function
    End If
    
    If rsTmp.EOF Then
        RetMsg = "SO001 查無 [公司別] 為 " & intComp & " , [客編] 為 " & lngCustId & " 的 [地址編號] 資料 !"
        Exit Function
    Else
        lngInstAddrNo = rsTmp!InstAddrNo & ""
    End If
    
    CloseRecordset rsTmp
    
    If lngInstAddrNo = 0 Then RetMsg = "裝機 [地址編號] 為 0 不正確 !": Exit Function
    
    GetSysPara
    
    strPrcType = CmdID
    
    If CmdSeqNo = "" Then
'       #3327 整合流程時 , 不做 IsDataOK
'        If Not IsDataOK(rs04, CmdID, FaciCode, FaciSno) Then
'            RetMsg = GetMsg(RetMsg)
'            Exit Function
'        End If
    Else
        GetRS307 , CmdSeqNo
        If rs307.State <= 0 Or rs307.EOF Then
            RetMsg = "SO307 查無 [序號] 為 " & CmdSeqNo & " 的資料 !"
            Exit Function
'        Else
'            With rs307
'                ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, _
'                                                !NewModemMac & "", _
'                                                !NewModelName & "", _
'                                                !NewCMBaudRate & "", _
'                                                Val(!NewDynIPCount & ""), _
'                                                Val(!NewFixIPCount & ""), _
'                                                !NewHFCNode & "", _
'                                                !AddCPEMac & "", !DelCPEMac & "", _
'                                                !AddCPEStatIP & "", !DelCPEStatIP & "", _
'                                                !NewIPaddress & "", _
'                                                !NewMTAMAC & "", , , , CmdSeqNo)
'            End With
'            RetMsg = GetMsg(RetMsg)
'            Exit Function
        End If
    End If
    
    ' C6 命令無 Update SO004 等 Table , 怪怪 , 確認中 ..
    Select Case CmdID
        
        ' 裝機,拆機,軟拆,軟復,CP Only 裝機,CP Only 拆機,CP裝分線,CP拆分線
        Case "C1", "E1", "C2", "E2", "A1_D", "A1_C", "E3", "A1_P", "E4", "E5"
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , , , , , , , , , , , , CmdSeqNo)
        
        ' 6. EMTA IP 異動
        ' 1.  參考號2時傳送C6 指令，舊的NODE 取用裝機地址資訊，新NODE 取用畫面輸入或移裝工單的新地址資訊
        ' 2.  參考號5時傳送E6 指令，舊的NODE 取用裝機地址資訊，新NODE 取用畫面輸入或移裝工單的新地址資訊
        ' 3.  未勾選時按照原機制傳送A4 指令
        Case "A4", "E6" ' EMTA IP 異動
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , , rs307!NewHFCNode & "", , , , , rs307!NewIPaddress & "", , , , , CmdSeqNo)
        
        Case "A6" ' CPE MAC 異動
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , , , rs307!AddCPEMac & "", , , , , , , , , CmdSeqNo)
            
        Case "A5" ' CPE IP 異動
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , , , , , rs307!AddCPEStatIP & "", , , , , , , CmdSeqNo)
            
        Case "A2" ' 申請固定IP
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , Val(rs307!NewFixIPCount & ""), , rs307!AddCPEMac & "", , rs307!AddCPEStatIP & "", , , , , , , CmdSeqNo)
            
        Case "A3" ' 申請固定IP
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , Val(rs307!NewFixIPCount & ""), , , rs307!DelCPEMac & "", , rs307!DelCPEStatIP & "", , , , , , CmdSeqNo)
            
        Case "A1_MIP", "A1_PIP" ' 取消動態IP , 申請動態IP
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , Val(rs307!NewDynIPCount & ""), , , , , , , , , , , , CmdSeqNo)
        
        Case "RP", "A7", "A8", "A9", "A10" ' CM / EMTA 更換
            If CmdSeqNo = Empty Then
                If Len(FaciCode) > 0 And Len(FaciSno) > 0 Then
                    strFC = FaciCode & ""
                    strFN = FaciName & ""
                    strFCsno = FaciSno
                    If CmdID = "RP" Then
                        ExecCmd = ParaClt(rs04, Get_RP_Cmd(rs04, FaciCode), , , , lngCustId, intComp, FaciSno, GetModelName(FaciSno), , , , , , , , , , , , , , CmdSeqNo)
                    Else
                        ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, FaciSno, GetModelName(FaciSno), , , , , , , , , , , , , , CmdSeqNo)
                    End If
                Else
                    RetMsg = "CM / EMTA 更換命令 , 請確認 [FaciSno] 及 [FaciCode] 參數有值 !"
                End If
            Else
                FaciSno = Replace(rs307!NewModemMac & "", ":", "", 1)
                strFCsno = FaciSno
                GetFaciName lngCustId
                ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, FaciSno, GetModelName(FaciSno), , , , , , , , , , , , , , CmdSeqNo)
            End If
        
        Case "A1_BU", "TRIAL", "STOPTRIAL"
                ' "STOPTRIAL" BaudRate = ""
                If CmdID <> "STOPTRIAL" Then
                    If BaudRate = Empty Then
                        If CmdSeqNo <> "" And rs307.State > 0 And Not rs307.EOF Then
                            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , rs307!NewCMBaudRate & "", , , , , , , , , , , , , CmdSeqNo)
                        Else
                            AddRetMsg "請傳入 [ 新速率 ] !"
                            ExecCmd = False
                        End If
                    Else
                        ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , Get_BaudRate_Desc(BaudRate), , , , , , , , , , , , , CmdSeqNo)
                    End If
                Else
                    ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , , , , , , , , , , , , CmdSeqNo)
                End If
        
        Case ""
            RetMsg = "請傳入命令代號 !"
        
        Case Else
            With rs307
                ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, _
                                                !NewModemMac & "", _
                                                !NewModelName & "", _
                                                !NewCMBaudRate & "", _
                                                Val(!NewDynIPCount & ""), _
                                                Val(!NewFixIPCount & ""), _
                                                !NewHFCNode & "", _
                                                !AddCPEMac & "", !DelCPEMac & "", _
                                                !AddCPEStatIP & "", !DelCPEStatIP & "", _
                                                !NewIPaddress & "", _
                                                !NewMTAMAC & "", , , , CmdSeqNo)
            End With
'            RetMsg = "命令代號錯誤 !"
        
'        Case "A1_BU"
'            If Len(BuadRate) > 0 Then
'                ExecCmd = ParaClt(rs04, "A1_BU", , , , lngCustId, intComp, , , BuadRate)
'            Else
'                RetMsg = "速率升降級命令 , 請確認 [新速率] 參數有值 !"
'            End If
        
    
    End Select
    
    If Not ExecCmd And ClearSNo Then
        objCmdCN.Execute "UPDATE " & strCMowner & "SO307 SET SNO=NULL" & _
                                        " WHERE CMDSEQNO='" & CmdSeqNo & "'"
    End If
    
    RetMsg = GetMsg(RetMsg)
    CmdSeqNo = strRetCmdSeqNo
  Exit Function
ChkErr:
    ErrSub "clsCommand", "ExecCmd"
End Function

'Private Function Get_BaudRate_Code(ByRef BaudRateDesc As String) As String
'  On Error GoTo ChkErr
'    strQry = "SELECT CODENO FROM " & GetOwner & "CD064 WHERE DESCRIPTION='" & BaudRateDesc & "'"
'    Get_BaudRate_Code = GetRsValue(strQry, gcnGi) & ""
'  Exit Function
'ChkErr:
'    ErrSub "clsCommand", "Get_BaudRate_Code"
'End Function



Private Function GetMsg(M As String) As String
  On Error GoTo ChkErr
    If M <> Empty Then
        If strRetMsg <> Empty Then M = M & vbCrLf & strRetMsg
    Else
        M = strRetMsg
    End If
    GetMsg = Replace(M, strCrLf, vbCrLf, 1)
  Exit Function
ChkErr:
    ErrSub "clsCommand", "GetMsg"
End Function

Private Function IsDataOK(ByRef rs04 As Object, CmdID As String, _
                                            Optional FaciCode As String = "", Optional FaciSno As String = "") As Boolean
                                            
  On Error GoTo ChkErr
  
    Dim strCmd As String
    IsDataOK = False
    
    If strResvTime <> "" Then
        If CDate(strResvTime) <= RightNow Then
            AddRetMsg "[預約時間] 需大於現在時間"
            Exit Function
        Else
'            AddRetMsg "設定預約時間, 將不等待 Gateway 回應 !"
        End If
    End If
    
    With rs04
    
        Select Case CmdID
        
            Case "E1", "C1" ' 1. 裝機
                ' 設備拆機日期是空白與開機日期 是空白 才可以裝機指令 此功能才可以ENABLE
                If !CMOpenDate & "" <> "" Or !PRdate & "" <> "" Then
                    AddRetMsg "設備 [開機日期] 與 [拆機日期] 無值才可裝機 !"
                    Exit Function
                End If
                
                Dim blnChkAddrNo As Boolean
                
                GetHFCNode !CustID, !CompCode, !ServiceType & "", blnChkAddrNo
                
                If blnChkAddrNo Then
                    AddRetMsg "該客戶移機中，未指定新址，請指定新址後再送開機指令！"
                    Exit Function
                End If
                
            Case "E2", "C2" ' 2. 拆機 ( 含裝機退件 )
                ' 設備裝機日期非空白 必須有開機日期才可以作拆機指令 此功能才可以ENABLE
                ' If !InstDate & "" = "" Or !CMOpenDate & "" = "" Then
                If !CMOpenDate & "" = "" Then
                    AddRetMsg "設備 [裝機日期] 與 [開機日期] 有值才可拆機 !"
                    Exit Function
                End If
                
            Case "A1_D" ' 3. 軟拆
                ' 有開機日 且開機日>(關機日OR 無關機日)則只可選軟關
                ' 正常使用中、安裝中不分EMTA 與CM 該筆資料必須（有開機日期 與 關機日期是空的） 或（開機日期>關機日期）此功能才可以ENABLE
                If !CMOpenDate & "" = "" Or !CMCloseDate & "" <> "" Then
                    If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                        If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                            AddRetMsg "設備 [正常使用中] 才可軟拆 !"
                            Exit Function
                        End If
                    Else
                        AddRetMsg "設備 [正常使用中] 才可軟拆 !"
                        Exit Function
                    End If
                End If
                
            Case "A1_C" ' 4. 軟復
                ' 軟關與軟復兩各選項是互斥，設備若(關機日OR 無關機日) > 開機日 則只可以選軟復
                ' 正常使用中、安裝中不分EMTA 與CM 該筆資料必須（有關機日期） 或（開機日期<關機日期）此功能才可以ENABLE
                If !CMOpenDate & "" <> "" Or !CMCloseDate & "" = "" Then
                    If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                        If CDT(!CMCloseDate) < CDT(!CMOpenDate) Then
                            AddRetMsg "設備 [正常使用中] 才可軟復 !"
                            Exit Function
                        End If
                    Else
                        AddRetMsg "設備 [正常使用中] 才可軟復 !"
                        Exit Function
                    End If
                End If
                
            Case "E3" ' 5. CP 裝機
                ' 設備正常使用中、安裝中EMTA 開機日期大於關機日期  且關聯到CP 服務需有CP 對應門號，且該門號為安裝中或正常此功能才可以ENABLE
                If Not Is_EMTA_Dev(rs04) Then
                    AddRetMsg "此設備非 EMTA , 無法進行 CP 裝機 !"
                    Exit Function
                End If
                
                ' #4302,2009/01/05 By Hammer,(RA)_SO1171A_20081223_調整CLI控制台CP門號判斷.doc
                If IsCP(rs04) Then
                    If Not IsCMCP(rs04) Then
                        AddRetMsg "無 [ 安裝中 ] 或 [ 正常 ] 的 [ CP 門號 ] , 無法進行 CP 裝機 !"
                        Exit Function
                    End If
                End If
'                IsCMCP = (GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                                        " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                                                        " AND COMPCODE=" & rsData("CompCode") & _
                                                        " AND CUSTID=" & rsData("CustID") & _
                                                        " AND SERVICETYPE='P'" & _
                                                        " AND PRDATE IS NULL") = 1)
                If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                    If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                        AddRetMsg "無 [ 安裝中 ] 或 [ 正常 ] 的 [ CP 門號 ] , 無法進行 CP 裝機 !"
                        Exit Function
                    End If
                End If
                ' 請加強資料搜尋時接以該客編相同的資料為主（custid = 設定的 custid），
                ' gateway 取值時請增加group by Gateway ，
                ' 如果取出有兩筆（含）以上時需要提出警示訊息（該emta 門號分別屬於不同 mediaGateway 無法設定），不做指令傳送
                If Val(GetRsValue("SELECT COUNT(GATEWAY) FROM " & GetOwner & "SO004" & _
                                    " WHERE REFACISNO='" & !FaciSno & "'" & _
                                    " AND CUSTID=" & !CustID & _
                                    " AND COMPCODE=" & !CompCode & _
                                    " AND SERVICETYPE='P'" & _
                                    " AND PRDATE IS NULL" & _
                                    " And FaciCode In " & _
                                    " (Select CodeNo From " & GetOwner & "CD022 Where Nvl(CPImportMode,0) In (0,1))" & _
                                    " GROUP BY GATEWAY", gcnGi, "") & "") > 1 Then
                    AddRetMsg "該 EMTA 門號分別屬於不同 Media Gateway 無法設定 !"
                    Exit Function
                End If
                
            Case "A1_P" ' 6. CP 拆機
                ' 設備正常使用中、安裝中 EMTA 開機日期大於關機日期 且關聯到P 服務 CP 對應門號必須要都是拆機的 此功能才可以ENABLE
                If Not Is_EMTA_Dev(rs04) Then
                    AddRetMsg "此設備非 EMTA , 無法進行 CP 拆機 !"
                    Exit Function
                End If
                If IsCP(rs04) Then
                    If Not IsCMCP(rs04) Then
                        AddRetMsg "無 [ 安裝中 ] 或 [ 正常 ] 的 [ CP 門號 ] , 無法進行 CP 拆機 !"
                        Exit Function
                    End If
                End If
                
            Case "RP" ' CM / EMTA 更換
                If Len(FaciCode) = 0 Then
                    AddRetMsg "[品名] 為必要欄位,須有值 !!"
                    Exit Function
                End If
                If Len(Trim(FaciSno)) = 0 Then
                    AddRetMsg "[CM 序號] 為必要欄位,須有值 !!"
                    Exit Function
                Else
                    ' 2006/06/21 輸入新序號時需要檢核該序號在SO004 中是否有重複(沒有拆機日期的)及算重複不可作設定
                    Dim strFaciSno As String
                    strFaciSno = FaciSno
                    strFaciSno = Replace(strFaciSno, "-", "", 1)
                    strFaciSno = Replace(strFaciSno, ":", "", 1)
                    strCmd = Get_RP_Cmd(rs04, FaciCode)
                    If strFaciSno = !FaciSno & "" Or strFaciSno = !ReFaciSno & "" Then
                        AddRetMsg "[CM 序號] 與舊序號相同 .. 請重新輸入新序號 !"
                        Exit Function
                    Else
                        ' Jacky 說控制台再判斷就好
                        ' " OR REFACISNO='" & strFaciSno & "' )"
'                        If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
'                                                " WHERE FACISNO='" & strFaciSno & "' AND PRDATE IS NULL") > 0 Then
'                            AddRetMsg "此序號已被使用 , 請重新輸入 !"
'                            Exit Function
'                        End If
                        If strCmd = "A8" Or strCmd = "A9" Then
                            If GetMTAMAC(strFaciSno) = Empty Then
                                AddRetMsg "此序號 [ " & strFaciSno & " ] 入庫資料找不到 !"
                                Exit Function
                            End If
                        End If
                    End If
                    
                    ' EMTA 換 CM , 要檢核 , 如果 ReFaciSno 有等於 , 則 Show Msg
                    If strCmd = "A10" Then
                        If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                                " WHERE REFACISNO='" & !FaciSno & "'" & _
                                                " AND PRDATE IS NULL AND SERVICETYPE='P'") > 0 Then
                            AddRetMsg "有 CP 門號 , 不可換成 CM "
                            Exit Function
                        End If
                    End If
                End If
                
            Case "E4" ' CP裝分線
                ' 正常使用中、安裝中 EMTA 開機日期大於關機日期 且關聯到P 服務 CP 對應門號必須要有一筆（含以上）正常，
                ' 新的一筆安裝中（無裝拆機日期） 此功能才可以ENABLE
                ' SNO <> "" AND InstDate is Null AND PRdate Is Null
                If Not Is_EMTA_Dev(rs04) Then
                    AddRetMsg "此設備非 EMTA , 無法進行 CP 裝分線 !"
                    Exit Function
                End If
                If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                    If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                        AddRetMsg "無 [ 安裝中 ] 或 [ 正常 ] 之 EMTA 設備 , 無法進行 CP 裝分線 !"
                        Exit Function
                    End If
                End If
                If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                " WHERE REFACISNO='" & !FaciSno & "'" & _
                                " AND COMPCODE=" & !CompCode & " AND CUSTID=" & !CustID & _
                                " AND SERVICETYPE='P' AND SNO IS NOT NULL " & _
                                " AND PRDATE IS NULL AND INSTDATE IS NULL") < 1 Then
                    AddRetMsg "此 EMTA 無對應 CP 門號 , 無法安裝 CP 分線 !"
                    ' MsgBox "此門號無正常對應之 EMTA 設備 , 無法安裝 CP 分線 !", vbInformation, "訊息"
                    Exit Function
                End If
                
            Case "E5" ' CP 拆分線
                ' 正常使用中、安裝中 EMTA 開機日期大於關機日期 且 關聯到P 服務 CP 對應門號必須要有一筆（含以上）正常，有一筆拆除中（有裝機日期、無拆機日期） 此功能才可以ENABLE
                ' PRSNO <> "" and PRdate =""
                If Not Is_EMTA_Dev(rs04) Then
                    AddRetMsg "此設備非 EMTA , 無法進行 CP 拆分線 !"
                    Exit Function
                End If
                ' 正常使用中、安裝中 EMTA 開機日期大於關機日期 且關聯到該客戶P 服務的CP 對應門號必須要有一筆（含以上）正常，
                ' 新的一筆安裝中（無裝拆機日期） 此功能才可以ENABLE
                ' EMTA 序號關聯到so004.servicetype='P' and prdate is null 的取出 p 服務的 so004.Gateway 取一筆即可與cp 相同取值原則
                ' If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                '     If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                '         ' MsgBox "無 [ 安裝中 ] 或 [ 正常 ] 之 EMTA 設備 , 無法進行 CP 拆分線 !", vbInformation, "訊息"
                '         MsgBox "設備已關機 , 無法進行 CP 拆分線 !", vbInformation, "訊息"
                '         Exit Function
                '     End If
                ' End If
                ' PRSNO  有值, InstDate 有 , PRdate 無
                ' AND SNO IS NOT NULL
                If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                " WHERE REFACISNO='" & !FaciSno & "'" & _
                                " AND COMPCODE=" & !CompCode & " AND CUSTID=" & !CustID & _
                                " AND SERVICETYPE='P'" & _
                                " AND PRDATE IS NULL AND INSTDATE IS NOT NULL") <= 1 Then
                    AddRetMsg "無法拆分機 !"
                    ' MsgBox "無正常的 CP 門號 , 無法 CP 拆分線 !", vbInformation, "訊息"
                    ' MsgBox "此門號無正常對應之 EMTA 設備 , 無法安裝 CP 拆分線 !", vbInformation, "訊息"
                    Exit Function
                End If
                ' 停拆移機單號
                If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                " WHERE REFACISNO='" & !FaciSno & "'" & _
                                " AND COMPCODE=" & !CompCode & " AND CUSTID=" & !CustID & _
                                " AND SERVICETYPE='P' AND PRSNO IS NOT NULL" & _
                                " AND PRDATE IS NULL AND INSTDATE IS NOT NULL") > 1 Then
                    AddRetMsg "有 2 張 (以上) 拆機單 , 請下拆機命令 !"
                    Exit Function
                End If
                
            Case "A1_BU", "TRIAL", "STOPTRIAL"
            
                    If strBaudRate = Empty Or Get_BaudRate_Desc(strBaudRate) & "" = !CMBaudRate & "" Then
                        AddRetMsg "請傳入 [ 新速率 ] !"
                        Exit Function
                    End If
                    
                    If CmdID = "TRIAL" Then
                        If strProbStopDate = Empty Then
                            AddRetMsg "請傳入 [ 預計試用截止日 ] 後再進行 [ 設定 ] !"
                            Exit Function
                        End If
                        ' 預計試用截止日 : ProbationStopDate
                        ' 停止試用狀態 : ProbationStopFlag ( 0=否, 1=是 )
                        ' 且(SO004.ProbationStopDate無值或SO004. ProbationStopFlag=1)才能下此命令。
                        If !ProbationStopDate & "" = Empty Or !ProbationStopFlag = 1 Then
                        Else
                            AddRetMsg "無法下 [試用速率] 命令 !"
                            Exit Function
                        End If
                    End If
                    
                    If CmdID = "STOPTRIAL" Then
                        If !ProbationStopDate & "" <> Empty And !ProbationStopFlag = 0 Then
                        Else
                            AddRetMsg "無法下 [終止試用] 命令 !"
                            Exit Function
                        End If
                    End If
                    ' 安裝中或正常的設備
                    If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                        If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                            If CmdID = "A1_BU" Then
                                AddRetMsg "無 [ 安裝中 ] 或 [ 正常 ] 之設備 , 無法進行 [速率升降級] !"
                            ElseIf CmdID = "TRIAL" Then
                                AddRetMsg "無 [ 安裝中 ] 或 [ 正常 ] 之設備 , 無法下 [試用速率] 命令 !"
                            Else
                                AddRetMsg "無 [ 安裝中 ] 或 [ 正常 ] 之設備 , 無法下 [終止試用] 命令 !"
                            End If
                            Exit Function
                        End If
                    End If
                    
        End Select
        
    End With
    
    IsDataOK = True
    
  Exit Function
ChkErr:
    ErrSub "clsCommand", "IsDataOK"
End Function

Private Sub GetFaciName(lngCustId As Long)
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    strQry = "SELECT FACICODE,FACINAME FROM " & GetOwner & "SO004" & _
                    " WHERE CUSTID=" & lngCustId & _
                    " AND COMPCODE=" & gCompCode & _
                    " AND FACISNO='" & strFCsno & "'"
    If GetRS(rsTmp, strQry, gcnGi) Then
        If rsTmp.State > 0 And Not rsTmp.EOF Then
            strFC = rsTmp!FaciCode & ""
            strFN = rsTmp!FaciName & ""
        End If
    End If
    On Error Resume Next
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
    ErrSub "clsCommand", "GetFaciName"
End Sub

'參考號2 換新CM A7
'參考號2 舊CM / 換新參考號5 EMTA A9
'參考號5 換新EMTA A8
'參考號5 EMTA / 換新參考號2 CM A10
Private Function Get_RP_Cmd(ByRef rs04 As Object, FaciCode As String) As String
  On Error GoTo ChkErr
    
    Dim NewDevRefNo As Integer
    
    strQry = "SELECT REFNO FROM " & GetOwner & "CD022 WHERE CODENO=" & FaciCode
    NewDevRefNo = Val(GetRsValue(strQry, gcnGi))
    
    strQry = "SELECT REFNO FROM " & GetOwner & "CD022 WHERE CODENO=" & rs04("FaciCode")
    
    Select Case Val(GetRsValue(strQry, gcnGi))
        Case 2
            Get_RP_Cmd = IIf(NewDevRefNo = 2, "A7", "A9")
        Case 5
            Get_RP_Cmd = IIf(NewDevRefNo = 5, "A8", "A10")
    End Select
    
  Exit Function
ChkErr:
    ErrSub "clsCommand", "Get_RP_Cmd"
End Function

Private Function Is_EMTA_Dev(ByRef rs04 As Object) As Boolean
  On Error GoTo ChkErr
    If rs04.State > 0 Then
        If rs04("FaciCode") & "" = Empty Then
            AddRetMsg "注意 [客戶設備資料檔].[品名代碼] 為空!"
        Else
            Is_EMTA_Dev = Val(GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD022" & _
                                            " WHERE CODENO=" & rs04("FaciCode") & _
                                            " AND REFNO=5", gcnGi, "") & "") > 0
        End If
    End If
  Exit Function
ChkErr:
    ErrSub "clsCommand", "Is_EMTA_Dev"
End Function

' CD022 品名編號代碼檔
' CPImportMode 門號匯入模式
' 0-台固、1-亞太、2-速博
' 0,1 為 CP , 2 為 VOIP

Private Function IsCMCP(ByRef rs04 As Object) As Boolean
  On Error GoTo ChkErr
        IsCMCP = GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                            " WHERE REFACISNO='" & rs04("FaciSno") & "'" & _
                            " AND COMPCODE=" & rs04("CompCode") & _
                            " AND CUSTID=" & rs04("CustID") & _
                            " AND SERVICETYPE='P'" & _
                            " AND PRDATE IS NULL") > 1
  Exit Function
ChkErr:
    ErrSub "clsCommand", "IsCMCP"
End Function

