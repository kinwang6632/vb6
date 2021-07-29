Attribute VB_Name = "mod_ControlPanel"
Option Explicit

Private Const FrmName = "mod_ControlPanel"

Public objCmdCN As Object

Public gCompFilterStr As String
Public gCompCode As String
Public gServiceType As String
Public gDefaultComp As String
Public objActForm As Form
Public gLogInID As String
Public gstrUserPremission As String
Public int_EMTA_IP_Type As Integer
Public intTimeOut As Integer
Public strCMowner As String
Public strResvTime As String
Public lngInstAddrNo As Long
Public strPrcType As String

Public Const Alfa = False
Public mFlds As GiGridFlds
Public Const LinkOdie = False
Public strErrMsg As String

Public strMC As String ' ModelCode
Public strMN As String ' ModelName
Public strFC As String ' FaciCode
Public strFN As String  ' FaciName
Public strFCsno As String ' FaciSno
Public strSno As String
Public strMediaBillNo As String
Public strBaudRate As String
Public strNewBaudRate As String
Public strNewModelCode As String
Public strNewFaciSno As String

Public strCmdDIY As String
Public strProbStopDate As String

'Public strIPnum As String

Public strDynaIPcnt As String
Public strFixIPcnt As String

Public strIP As String
Public strCPE As String

Public strOldIP As String
Public strOldCPE As String
Public strNode As String
Public strNewIP As String


Public blnSendKey As Boolean
Public blnShowFaci As Boolean

Private Const blnDebug = True
Public rs307 As New ADODB.Recordset
Private rsResult As New ADODB.Recordset
Public rsUpd As New ADODB.Recordset
Public strRetCmdSeqNo As String
Public strFaciAuthorityField As String
Private strCmdSno As String
Private strStatus As String
Private lngReturn As Long
Private strRetMsg As String
Private strFailureCode As String
Private strCmdName As String
Private blnNoShwMsg As Boolean
Private blnTimeOut As Boolean

Public Function CmdDisp(Value As Variant) As Variant
  On Error GoTo ChkErr
    '#6776 改用case 的方式,並增加RP By Kin 2014/04/21
    Select Case UCase(Value)
        Case "C1"
            CmdDisp = "裝機 (CM)"
        Case "E1"
            CmdDisp = "裝機 (EMTA)"
        Case "C2"
            CmdDisp = "拆機 (CM裝機退件)"
        Case "E2"
            CmdDisp = "拆機 (EMTA裝機退件)"
        Case "E3"
            CmdDisp = "CP 裝機"
        Case "A1_P"
            CmdDisp = "CP Only 拆機"
        Case "E4"
            CmdDisp = "EMTA 加裝 CP 門號"
        Case "E5"
            CmdDisp = "EMTA 拆 CP 門號"
        Case "A1_D"
            CmdDisp = "軟拆"
        Case "A1_C"
            CmdDisp = "軟復"
        Case "A1_BU"
            CmdDisp = "速率升降級"
        Case "A1_PIP"
            CmdDisp = "申請動態IP"
        Case "A1_MIP"
            CmdDisp = "取消動態IP"
        Case "Q1"
            CmdDisp = "CM reset"
        Case "Q2"
            CmdDisp = "查詢CM狀況"
        Case "A7"
            CmdDisp = "CM 拆舊 / 換新"
        Case "A8"
            CmdDisp = "EMTA 拆舊 / 換新"
        Case "A9"
            CmdDisp = "拆舊CM / 換新EMTA"
        Case "A10"
            CmdDisp = "拆舊EMTA / 換新 CM"
        Case "A6"
            CmdDisp = "變更 CPE MAC"
        Case "A5"
            CmdDisp = "CM IP  CPE IP 設定"
        Case "A4"
            CmdDisp = "EMTA IP Priv IP 設定"
        Case "A3"
            CmdDisp = "取消固定IP"
        Case "A2"
            CmdDisp = "申請固定IP"
        Case "C6"
            CmdDisp = "CM 移機"
        Case "E6"
            CmdDisp = "EMTA 移機"
        Case "TRIAL"
            CmdDisp = "試用速率"
        Case "STOPTRIAL"
            CmdDisp = "停止試用"
        Case "RP"
            CmdDisp = "設備換裝"
        Case Else
            CmdDisp = Value
    End Select
  
'    CmdDisp = Switch(Value = "C1", "裝機 (CM)", Value = "E1", "裝機 (EMTA)", _
'                                    Value = "C2", "拆機 (CM裝機退件)", Value = "E2", "拆機 (EMTA裝機退件)", _
'                                    Value = "E3", "CP 裝機", Value = "A1_P", "CP Only 拆機", _
'                                    Value = "E4", "EMTA 加裝 CP 門號", Value = "E5", "EMTA 拆 CP 門號", _
'                                    Value = "A1_D", "軟拆", Value = "A1_C", "軟復", _
'                                    Value = "A1_BU", "速率升降級", Value = "A1_PIP", "申請動態IP", _
'                                    Value = "A1_MIP", "取消動態IP", Value = "Q1", "CM reset", _
'                                    Value = "Q2", "查詢CM狀況", Value = "A7", "CM 拆舊 / 換新", _
'                                    Value = "A8", "EMTA 拆舊 / 換新", Value = "A9", "拆舊CM / 換新EMTA", _
'                                    Value = "A10", "拆舊EMTA / 換新 CM", Value = "A6", "變更 CPE MAC", _
'                                    Value = "A5", "CM IP  CPE IP 設定", Value = "A4", "EMTA IP Priv IP 設定", _
'                                    Value = "A3", "取消固定IP", Value = "A2", "申請固定IP", _
'                                    Value = "C6", "CM 移機", Value = "E6", "EMTA 移機", _
'                                    Value = "TRIAL", "試用速率", Value = "STOPTRIAL", "停止試用")
  Exit Function
ChkErr:
    ErrSub FrmName, "CmdDisp"
End Function

'CD022 品名編號代碼檔
'   1=解碼器    2=Cable Modem   3=STB   4=ICC
'   5=EMTA  6=CP    7=大樓設備  8=大樓設備孔位

Public Function ParaClt(ByRef rs As Object, ByVal strCmd As String, Optional ByVal strSeqNo As String = "", _
                                        Optional blnQuery As Boolean = False, Optional blnBatch As Boolean = False, _
                                        Optional lngCustId As Long = 0, Optional intComp As Integer = 0, _
                                        Optional strNewCMmac As String = "", Optional strNewModel As String = "", _
                                        Optional strNewBaudRateName As String = "", _
                                        Optional intNewDynIPcnt As Integer = 0, Optional intNewFixIPcnt As Integer = 0, _
                                        Optional strNewHFCNode As String = "", Optional strAddCPEMAC As String = "", _
                                        Optional strDelCPEMAC As String = "", Optional strAddCPEStatIP As String = "", _
                                        Optional strDelCPEStatIP As String = "", Optional strNewIPAddress As String = "", _
                                        Optional strNewMTAMAC As String = "", Optional strCmdStatus As String = "W", _
                                        Optional strExecEntry As String = "", Optional strExecTime As String = "", _
                                        Optional ByRef CmdSeqNo As String = "") As Boolean
  On Error GoTo ChkErr
  
    Dim strPriority As String, strSource As String
    Dim strMsgCode As String, strMsgName As String
    Dim strCMmac As String, strModel As String, strBaudRateName As String
    Dim intDynIPcnt As Integer, intFixIPcnt As Integer
    Dim strMediaGateway As String, strCPnum As String, strTel As String
    Dim strHFCNode As String, strCPEMac As String, strCPEStatIP As String
    Dim strIPAddress As String, strMTAMAC As String
    Dim blnChkAddrNo As Boolean
    
    With rs
        If strSeqNo <> Empty Then
            strCmdName = !MSGNAME & ""
            ParaClt = Ins2SwapTbl(rs, blnQuery, CmdSeqNo, strSeqNo, !Priority, !Source & "", !MsgCode & "", !MSGNAME & "", _
                                                    !CompCode, !CustID, !Modemmac & "", !NewModemMac & "", _
                                                    !ModelName & "", !NewModelName & "", !CMBaudRate, _
                                                    !NewCMBaudRate & "", Val(!DYNIPCOUNT), Val(!NewDynIPCount), _
                                                    Val(!FixIPCount), Val(!NewFixIPCount), !MediaGateway & "", _
                                                    !CPNumber & "", !Tel & "", !HFCNode & "", !NewHFCNode & "", _
                                                    !CPEMAC & "", !AddCPEMac & "", !DelCPEMac & "", _
                                                    !CPEStatIP & "", !AddCPEStatIP & "", !DelCPEStatIP & "", _
                                                    !IPADDRESS & "", !NewIPaddress & "", !MTAMAC & "", !NewMTAMAC & "", _
                                                    strCmdStatus, strExecEntry, strExecTime, rs)
            
        
        Else
        
            strPriority = IIf(blnBatch, "10", "00")
            strSource = IIf(blnBatch, "B", "C")
            strMsgCode = ""
            strMsgName = strCmd
            
            strCmdName = CmdDisp(strCmd)
'            strCmdName = Switch(strCmd = "C1" Or strCmd = "E1", "裝機", strCmd = "C2" Or strCmd = "E2", "拆機", _
                                                strCmd = "E3", "CP 裝機", strCmd = "A1_P", "CP 拆機", strCmd = "E4", "CP 裝分線", _
                                                strCmd = "E5", "CP 拆分線", strCmd = "A1_D", "CM 軟拆", strCmd = "A1_C", "CM 軟復", _
                                                strCmd = "A1_BU", "速率升降級", strCmd = "A1_PIP", "申請動態IP", _
                                                strCmd = "A1_MIP", "取消動態IP", strCmd = "Q1", "重設 CM reset", strCmd = "Q2", "查詢CM狀況", _
                                                strCmd = "A7" Or strCmd = "A8" Or strCmd = "A9" Or strCmd = "A10", "CM / EMTA 更換", _
                                                strCmd = "A6", "CPE MAC 異動", strCmd = "A5", "CPE IP 異動", strCmd = "A4", "EMTA IP異動", _
                                                strCmd = "A3", "取消固定IP", strCmd = "A2", "申請固定IP")
            
            Select Case strMsgName
                        Case "C1", "E1", "C2", "E2", "E3", "A1_P", "E4", "E5", "A1_D", "A1_C"
                                strNewCMmac = !FaciSno & ""
                                strNewModel = GetModel(!ModelCode & "")
                                strNewBaudRateName = !CMBaudRate & ""
                                intNewDynIPcnt = Val(!DYNIPCOUNT & "")
                                intNewFixIPcnt = Val(!FixIPCount & "")
                                strNewHFCNode = GetHFCNode(lngCustId, intComp, !ServiceType & "", blnChkAddrNo)
                                strAddCPEMAC = GetCPEmac(!SeqNo & "", lngCustId)
                                strAddCPEStatIP = GetIPaddr(!SeqNo & "", lngCustId)
                                strNewIPAddress = !IPADDRESS & ""
'                                strNewMTAMAC = GetMTAMAC(!FaciSno & "")

                                If (strMsgName = "C1" Or strMsgName = "E1") And blnChkAddrNo Then
                                    If UI_Mode Then
                                        MsgBox "該客戶移機中，未指定新址，請指定新址後再送開機指令！", vbInformation, "訊息"
                                    Else
                                        AddRetMsg "該客戶移機中，未指定新址，請指定新址後再送開機指令！"
                                    End If
                                    ParaClt = False
                                    Exit Function
                                End If

                        Case Else
                                strCMmac = !FaciSno & ""
                                strModel = GetModel(!ModelCode & "")
                                If strCmd = "STOPTRIAL" Then
                                    strBaudRateName = !ProbationRateName & ""
                                    strNewBaudRateName = !CMBaudRate & ""
                                Else
                                    strBaudRateName = !CMBaudRate & ""
                                End If
                                intDynIPcnt = Val(!DYNIPCOUNT & "")
                                intFixIPcnt = Val(!FixIPCount & "")
                                strHFCNode = GetHFCNode(lngCustId, intComp, !ServiceType & "")
                                If strMsgName = "A6" Then
                                    If CmdSeqNo = "" Then
                                        strCPEMac = Screen.ActiveForm.cboCPE.Text
                                    Else
                                        strCPEMac = rs307!AddCPEMac & ""
                                    End If
                                Else
                                    If CmdSeqNo = "" Then
                                        strCPEMac = GetCPEmac(!SeqNo & "", lngCustId)
                                    Else
                                        If rs307.State > 0 And Not rs307.EOF Then
                                            strCPEMac = rs307!AddCPEMac & ""
                                        End If
                                    End If
                                End If
                                If strMsgName = "A5" Then
                                    If CmdSeqNo = "" Then
                                        strCPEStatIP = Screen.ActiveForm.cboIP.Text
                                    Else
                                        If rs307.State > 0 And Not rs307.EOF Then
                                            strCPEStatIP = rs307!AddCPEStatIP & ""
                                        End If
                                    End If
                                Else
                                    If CmdSeqNo = "" Then
                                        strCPEStatIP = GetIPaddr(!SeqNo & "", lngCustId)
                                    Else
                                        If rs307.State > 0 And Not rs307.EOF Then
                                            strCPEStatIP = rs307!AddCPEStatIP & ""
                                        End If
                                    End If
                                End If
                                strIPAddress = !IPADDRESS & ""
'                                strMTAMAC = GetMTAMAC(!FaciSno & "")
                                If strNewCMmac = Empty Then strNewCMmac = strCMmac
                                If strNewModel = Empty Then strNewModel = strModel
                                If strNewBaudRateName = Empty Then strNewBaudRateName = strBaudRateName
                                If intNewDynIPcnt = 0 Then intNewDynIPcnt = intDynIPcnt
                                If intNewFixIPcnt = 0 Then intNewFixIPcnt = intFixIPcnt
                                If strNewHFCNode = Empty Then strNewHFCNode = strHFCNode
                                If strAddCPEMAC = Empty Then strAddCPEMAC = strCPEMac
                                If strAddCPEStatIP = Empty Then strAddCPEStatIP = strCPEStatIP
                                If strNewIPAddress = Empty Then
                                    If strMsgName <> "C6" And strMsgName <> "E6" Then strNewIPAddress = strIPAddress
                                End If
'                                If strNewMTAMAC = Empty Then strNewMTAMAC = strMTAMAC
            End Select
            
            strCMmac = Mask(strCMmac)
            strNewCMmac = Mask(strNewCMmac)
            If strCMmac = Empty And strNewCMmac <> Empty Then strCMmac = strNewCMmac

'            If strMTAMAC = Empty And strNewMTAMAC <> Empty Then strMTAMAC = strNewMTAMAC
            ' Jim 說 : MTAMAC , NEWMTAMAC 要填 CMMAC , NEWCMMAC
            strMTAMAC = strCMmac
            strNewMTAMAC = strNewCMmac
            
            If strIPAddress = Empty And strNewIPAddress <> Empty Then strIPAddress = strNewIPAddress
            If strHFCNode = Empty And strNewHFCNode <> Empty Then strHFCNode = strNewHFCNode
             
            If strNewHFCNode = Empty And strHFCNode <> Empty Then strNewHFCNode = strHFCNode
            
            If strCPEMac <> Empty Then
                strCPEMac = Replace(strCPEMac, "-", ":", 1)
                strCPEMac = Mask(strCPEMac)
            End If
            If strAddCPEMAC <> Empty Then
                strAddCPEMAC = Replace(strAddCPEMAC, "-", ":", 1)
                strAddCPEMAC = Mask(strAddCPEMAC)
            End If
            If strDelCPEMAC <> Empty Then
                strDelCPEMAC = Replace(strDelCPEMAC, "-", ":", 1)
                strDelCPEMAC = Mask(strDelCPEMAC)
            End If
            
            strMediaGateway = GetMediaGw(!FaciSno & "", !CustID)
                    
            ' MTAMAC , NEWMTAMAC 填 FaciSno
            If Alfa And strMediaGateway = Empty And (strCmd = "A1_P" Or strCmd = "A1_BU" Or strCmd = "TRIAL") Then
                strMediaGateway = "BSTG01"
            End If

            If strCmd = "E5" Then
                strCPnum = GetCPno(!FaciSno & "", !CustID, True)
            Else
                strCPnum = GetCPno(!FaciSno & "", !CustID)
            End If
            
            strTel = !ContTel & ""
            
            If Alfa And (strCmd = "E4" Or strCmd = "E5") Then
                If strCPnum = Empty Then strCPnum = "000CE5BAF594"
                If strTel = Empty Then strTel = "22663388"
            End If
            
            If strCmd = "A9" Then
                ' 4288 需求調整如下:
                ' 原本CLI控制台CM換EMTA 是直接在控制台操作， 使用的方式是如果要由CM轉成EMTA 時，
                ' 在CLI控制台輸入新的EMTA序號，按下設定，程式會在A項設備自動新增一筆EMTA設備。
                ' 不需要派工單，不用自己去新增設備，更不用去指定EMTA上的CPIP。
                ' 透過CLI派工流程， CM換EMTA要由維換工單來派遣，在進CLI控制台前，設備已經新增好了，
                ' 在新增EMTA設備時CPIP也會由系統自動帶入，因此在進入CLI控制台送指令時，就不需要再去取一次CPIP。
                ' 我想問題的解決方式，應該是判斷送指令時，EMTA上的CPIP是否有值，
                ' 如果有值請以EMTA上的CPIP為主，不需要再重新取值，
                ' 如果沒有，請依照原本的規則自動取一筆CPIP給EMTA。
                strNewIPAddress = GetRsValue("Select IPAddress" & _
                                                " From " & GetOwner & "SO004" & _
                                                " Where Custid=" & rs("Custid") & _
                                                " And FaciSno='" & Replace(strNewCMmac, ":", "") & "'") & ""
                If strNewIPAddress = Empty Then strNewIPAddress = GetNewIP ' #2688 , 參考號2舊CM /換新參考號5 EMTA A9 需要自動依據NODE 取新的IPADDRESS
            End If
            
            '#6066 速率升降級如果有傳入新設備,要以新設備為主 By Kin 2011/07/26
            If UCase(strCmd) = UCase("A1_BU") Then
               
                'NewCMBaudRate
                'NewCPEIPRule  沒用到
                'NewDynIPCount
                'NewFixIPCount
                'NewHFCNode
                'NewIPaddress
                'NewModelName
                'NewModemMAC
                'NewMTAMAC
                
                'strNewBaudRateName = rsUpd("CMBaudRate")
                If strNewBaudRate <> Empty Then
                    strNewBaudRateName = Get_BaudRate_Desc(strNewBaudRate)
                End If
                If strDynaIPcnt <> Empty Then
                    intDynIPcnt = Val(strDynaIPcnt & "")
                End If
                If strFixIPcnt <> Empty Then
                    intFixIPcnt = Val(strFixIPcnt & "")
                End If
                'intDynIPcnt = Val(rsUpd("DynIPCount") & "")
                'intFixIPcnt = Val(rsUpd("FixIPCount") & "")
                If strNewIP <> Empty Then
                    strNewIPAddress = strNewIP
                End If
                'strNewIPaddress = rsUpd("IPAddress") & ""
                If strNewModelCode <> Empty Then
                    strNewModel = GetModel(strNewModelCode)
                End If
                If strNewFaciSno <> Empty Then
                    strNewCMmac = Mask(strNewFaciSno)
                    strNewMTAMAC = strNewCMmac
                End If
                
                     
            End If
            
            ParaClt = Ins2SwapTbl(rs, blnQuery, CmdSeqNo, strSeqNo, strPriority, strSource, strMsgCode, strMsgName, _
                                    intComp, lngCustId, strCMmac, strNewCMmac, strModel, strNewModel, _
                                    strBaudRateName, strNewBaudRateName, intDynIPcnt, intNewDynIPcnt, _
                                    intFixIPcnt, intNewFixIPcnt, strMediaGateway, strCPnum, strTel, _
                                    strHFCNode, strNewHFCNode, strCPEMac, strAddCPEMAC, strDelCPEMAC, _
                                    strCPEStatIP, strAddCPEStatIP, strDelCPEStatIP, strIPAddress, strNewIPAddress, _
                                    strMTAMAC, strNewMTAMAC, strCmdStatus, strExecEntry, strExecTime, rs)
        End If
    End With
    
    If CmdSeqNo = "" Then CmdSeqNo = strCmdSno
    
  Exit Function
ChkErr:
    ErrSub FrmName, "ParaClt"
End Function

Private Function Ins2SwapTbl(ByRef rs As Object, blnQuery As Boolean, ByRef CmdSeqNo As String, ParamArray varPara() As Variant) As Boolean
  On Error GoTo ChkErr
    
    Dim strTrcPos As String
    Dim varAry As Variant, varAry2 As Variant
    Dim bytLoop As Byte
    Dim strFaciSno As String, strSeqNo As String
    Dim strUpdSQL As String
    
    If Not Screen.ActiveForm Is Nothing Then
        If Not Screen.ActiveForm.rsData.EOF Then
            ' from console panel
            strFaciSno = Screen.ActiveForm.rsData!FaciSno & ""
            strSeqNo = Screen.ActiveForm.rsData!SeqNo & ""
        Else
            strFaciSno = rs!FaciSno & ""
            strSeqNo = rs!SeqNo & ""
        End If
    Else
        strFaciSno = rs!FaciSno & ""
        strSeqNo = rs!SeqNo & ""
    End If

'   strFaciSno = Replace(varPara(7), ":", "")
'   varPara(35)!Seqno
    varPara(4) = UCase(varPara(4))
    Debug.Print strSeqNo
    
    If CmdSeqNo <> Empty Then
        'GetRS307 , CmdSeqNo
        GoTo ReadyGoUpd
        Exit Function
    End If
    
    If blnQuery Then ' 查詢
        GetRS307 , CStr(varPara(0))
        GoTo Qry
    Else
        strCmdSno = Get_Cmd_Sno
        strRetCmdSeqNo = strCmdSno
        'CmdSeqNo = strCmdSno
        ' 重送 Chk_Resend
        If varPara(0) <> Empty Then
            strUpdSQL = "UPDATE " & strCMowner & "SO307" & _
                                    " SET CMDRESEND=" & strCmdSno & _
                                    " WHERE CMDSEQNO='" & varPara(0) & "'"
            objCmdCN.Execute strUpdSQL
        End If
    End If
    
    Select Case varPara(4)
        Case "A1_BU"
            strCmdDIY = "Free"
        Case "TRIAL"
            strCmdDIY = "Try"
            varPara(4) = "A1_BU"
        Case "STOPTRIAL"
            strCmdDIY = "Stop"
            varPara(4) = "A1_BU"
        Case Else
            strCmdDIY = ""
    End Select
    
    If GetRS307 Then
        With rs307
            If .State > 0 Then
                strTrcPos = 0
                .AddNew
                strTrcPos = "CmdSeqNo" & strCmdSno
                .Fields("CMDSEQNO") = strCmdSno
                strTrcPos = 1 & vbTab & varPara(1)
                .Fields("PRIORITY") = varPara(1)
                strTrcPos = 2 & vbTab & varPara(2)
                .Fields("SOURCE") = varPara(2)
                strTrcPos = 3 & vbTab & varPara(3)
                .Fields("MSGCODE") = IIf(varPara(3) = Empty, Null, varPara(3))
                strTrcPos = 4 & vbTab & varPara(4)
                .Fields("MSGNAME") = varPara(4)
                strTrcPos = 5 & vbTab & varPara(5)
                .Fields("COMPCODE") = varPara(5)
                strTrcPos = 6 & vbTab & varPara(6)
                .Fields("CUSTID") = varPara(6)
                strTrcPos = 7 & vbTab & varPara(7)
                .Fields("MODEMMAC") = varPara(7)
                strTrcPos = 8 & vbTab & varPara(8)
                .Fields("NEWMODEMMAC") = varPara(8)
                strTrcPos = 9 & vbTab & varPara(9)
                .Fields("MODELNAME") = varPara(9)
                strTrcPos = 10 & vbTab & varPara(10)
                .Fields("NEWMODELNAME") = varPara(10)
                strTrcPos = 11 & vbTab & varPara(11)
                .Fields("CMBAUDRATE") = varPara(11)
                strTrcPos = 12 & vbTab & varPara(12)
                .Fields("NEWCMBAUDRATE") = varPara(12)
                strTrcPos = 13 & vbTab & varPara(13)
                .Fields("DYNIPCOUNT") = varPara(13)
                strTrcPos = 14 & vbTab & varPara(14)
                .Fields("NEWDYNIPCOUNT") = varPara(14)
                strTrcPos = 15 & vbTab & varPara(15)
                .Fields("FIXIPCOUNT") = varPara(15)
                strTrcPos = 16 & vbTab & varPara(16)
                .Fields("NEWFIXIPCOUNT") = varPara(16)
                strTrcPos = 17 & vbTab & varPara(17)
                .Fields("MEDIAGATEWAY") = varPara(17)
                strTrcPos = 18 & vbTab & varPara(18)
                .Fields("CPNUMBER") = varPara(18)
                strTrcPos = 19 & vbTab & varPara(19)
                .Fields("TEL") = varPara(19)
                strTrcPos = 20 & vbTab & varPara(20)
                .Fields("HFCNODE") = varPara(20)
                strTrcPos = 21 & vbTab & varPara(21)
                .Fields("NEWHFCNODE") = varPara(21)
                strTrcPos = 22 & vbTab & varPara(22)
                .Fields("CPEMAC") = varPara(22)
                strTrcPos = 23 & vbTab & varPara(23)
                .Fields("ADDCPEMAC") = varPara(23)
                strTrcPos = 24 & vbTab & varPara(24)
                .Fields("DELCPEMAC") = varPara(24)
                strTrcPos = 25 & vbTab & varPara(25)
                .Fields("CPESTATIP") = varPara(25)
                strTrcPos = 26 & vbTab & varPara(26)
                .Fields("ADDCPESTATIP") = varPara(26)
                strTrcPos = 27 & vbTab & varPara(27)
                .Fields("DELCPESTATIP") = varPara(27)
                strTrcPos = 28 & vbTab & varPara(28)
                .Fields("IPADDRESS") = varPara(28)
                strTrcPos = 29 & vbTab & varPara(29)
                .Fields("NEWIPADDRESS") = varPara(29)
                strTrcPos = 30 & vbTab & varPara(30)
                .Fields("MTAMAC") = varPara(30)
                strTrcPos = 31 & vbTab & varPara(31)
                .Fields("NEWMTAMAC") = varPara(31)
                strTrcPos = 32 & vbTab & varPara(32)
                .Fields("COMMANDSTATUS") = varPara(32)
                strTrcPos = 33 & vbTab & varPara(33)
                .Fields("EXECENTRY") = IIf(varPara(33) = Empty, garyGi(1), varPara(33))
                strTrcPos = 34 & vbTab & varPara(34)
                .Fields("EXECTIME") = IIf(varPara(34) = Empty, RightNow, varPara(34))
                strTrcPos = "HostName" & vbTab & Get_Computer_Name
                .Fields("HOSTNAME") = Get_Computer_Name
                
                If strResvTime <> "" Then .Fields("ResvTime") = strResvTime
                
                If strSno <> "" Then .Fields("SNO") = strSno
                
                If varPara(4) = "A1_BU" Or varPara(4) = "TRIAL" Or varPara(4) = "STOPTRIAL" Then
                    If strSno <> "" Then
                        .Fields("CMBaudRateType") = 0 ' 0=派工升降
                    Else
                        Select Case strCmdDIY
                            Case "Free"
                                .Fields("CMBaudRateType") = 1 ' 1=免費升降
                            Case "Try"
                                .Fields("CMBaudRateType") = 2 ' 2=試用
                                .Fields("ProbationStopDate") = strProbStopDate
                            Case "Stop"
                                .Fields("CMBaudRateType") = 3 ' 3=停止試用
                                .Fields("ProbationStopDate") = RightDate
                            Case Else
                        End Select
                    End If
                End If
                
                ' QQ : 3358
                ' Date : 2008/01/28
                ' Doc : (SAD)3358_20070731_Lawrence.doc
                ' i.  速率升降級調整：增加回寫SO307.CMBaudRateType=1值 (1=免費升降)。
                ' CMBaudRateType : 0=派工升降, 1=免費升降, 2=試用, 3=停止試用
                
                If strMediaBillNo <> "" Then .Fields("MediaBillNo") = strMediaBillNo
                
                strTrcPos = "Update"
                .Update
                strTrcPos = "307 Update OK !"
                If Len(Dir("C:\Debug.txt")) > 0 Then
                     On Error Resume Next
                     Kill "C:\SO307.xml"
                    .Save "C:\SO307.xml", adPersistXML
                End If
                ' GetDTString
            End If
        End With
        
  On Error GoTo ChkErr
  
        If strResvTime <> Empty Then
            If Not Screen.ActiveForm Is Nothing Then ' from console panel
                With Screen.ActiveForm
                        GetRS307 , strCmdSno ' Requery data
                        .ShowResult rs307, .ggrQueryInfo ' Show Result
                        .pbr1.Max = 1
                        .pbr1.Value = 1
                End With
            End If
            If UI_Mode Then MsgBox "[ " & strCmdName & " ] 控制訊號已送出完成 !!  " & strCrLf & _
                                                    "Gateway 將於 [ 預約時間 ] : " & strResvTime & " 進行處理 !", vbInformation, "訊息"
            AddRetMsg "[ " & strCmdName & " ] 控制訊號已送出完成 !!  " & strCrLf & _
                                                    "Gateway 將於 [ 預約時間 ] : " & strResvTime & " 進行處理 !"
            Ins2SwapTbl = True
            Exit Function
        End If

    Dim pgb As ProgressBar
    If Screen.ActiveForm Is Nothing Then
        Set pgb = Nothing
    Else
        Set pgb = Screen.ActiveForm.pbr1
    End If
    
Qry:
        If Pooling(pgb, "Qry_Result") Then
            
            If UCase(strStatus) = "S" Then
                
ReadyGoUpd:
                
                Ins2SwapTbl = True
                strTrcPos = varPara(4)
                
                Select Case varPara(4)
                
                    Case "C1", "E1" ' 裝機
                        ' 指令成功回填開機日期
                        If strPrcType = "" Then
                             '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                             '判斷正常設備又要拿掉 For 銘誠 By Kin 2014/10/16
                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                    " SET CMOPENDATE=" & O2Date(RightNow) & _
                                                    ",UPDEN='" & garyGi(1) & "',UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                    " WHERE SEQNO='" & strSeqNo & "'"
                            
                            gcnGi.Execute strUpdSQL
                        Else
                            '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                            '判斷正常設備又要拿掉 For 銘誠 By Kin 2014/10/16
                            rs!CMOpenDate = RightNow
                        End If
                        
                    Case "C2", "E2"  ' 指令執行成功回填關機日期
                       
                        If strPrcType = "" Then
                            '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                            '判斷正常設備又要拿掉 For 銘誠 By Kin 2014/10/16
                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                    " SET CMCLOSEDATE=" & O2Date(RightNow) & _
                                                    ",UPDEN='" & garyGi(1) & "',UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                    " WHERE SEQNO='" & strSeqNo & "'"
                                                    
                                                    
                            gcnGi.Execute strUpdSQL
                        Else
                            '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                            '判斷正常設備又要拿掉 For 銘誠 By Kin 2014/10/16
                            rs!CMCloseDate = RightNow
                        End If
                        
                    Case "A1_C" ' CM 軟復
                        If strPrcType = "" Then
                            '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                            '判斷正常設備又要拿掉 For 銘誠 By Kin 2014/10/16
                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                    " SET CMOPENDATE=" & O2Date(RightNow) & _
                                                    ",UPDEN='" & garyGi(1) & "',UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                    " WHERE SEQNO='" & strSeqNo & "'"
                            gcnGi.Execute strUpdSQL
                        Else
                             '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                            '判斷正常設備又要拿掉 For 銘誠 By Kin 2014/10/16
                            rs!CMOpenDate = RightNow
                        End If
                        
                    Case "A1_D" ' CM 軟拆
                    
                        If strPrcType = "" Then
                            '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                            '判斷正常設備又要拿掉 For 銘誠 By Kin 2014/10/16
                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                    " SET CMCLOSEDATE=" & O2Date(RightNow) & _
                                                    ",UPDEN='" & garyGi(1) & "',UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                    " WHERE SEQNO='" & strSeqNo & "'"
                            gcnGi.Execute strUpdSQL
                        Else
                            '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                            rs!CMCloseDate = RightNow
                        End If
                        
                    Case "E3" ' CP裝機
                    
                        If strPrcType = "" Then
                              '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                    " SET CMOPENDATE=" & O2Date(RightNow) & _
                                                    ",UPDEN='" & garyGi(1) & "'" & _
                                                    ",UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                    " WHERE REFACISNO='" & strFaciSno & "'" & _
                                                    " AND COMPCODE=" & varPara(5) & _
                                                    " AND CUSTID=" & varPara(6) & _
                                                    " AND SERVICETYPE='P'" & _
                                                    " AND PRDATE IS NULL"
                                                        
                            gcnGi.Execute strUpdSQL
                        Else
                            '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                            '判斷正常設備又要拿掉 For 銘誠 By Kin 2014/10/16
                            rs!CMOpenDate = RightNow
                        End If
                                                    
'                        gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
                                                    " SET DYNIPCOUNT=" & varPara(14) & _
                                                    ",FIXIPCOUNT=" & varPara(16) & _
                                                    ",UPDEN='CABLESOFT'" & _
                                                    ",UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                    " WHERE FACISNO='" & strFaciSno & "'"
                                                    
                    Case "A1_P" ' CP拆機
                    
                        If strPrcType = "" Then
                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                    " SET CMCLOSEDATE=" & O2Date(RightNow) & _
                                                    ",UPDEN='" & garyGi(1) & "'" & _
                                                    ",UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                    " WHERE REFACISNO='" & strFaciSno & "'" & _
                                                    " AND COMPCODE=" & varPara(5) & _
                                                    " AND CUSTID=" & varPara(6) & _
                                                    " AND SERVICETYPE='P'" & _
                                                    " AND PRDATE IS NULL"
                            
                            gcnGi.Execute strUpdSQL
                        Else
                            '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                            '判斷正常設備又要拿掉 For 銘誠 By Kin 2014/10/16
                            rs!CMCloseDate = RightNow
                        End If
                    
                    Case "A1_BU", "TRIAL" ' 速率升降級
                        '#6066 如果有傳入RecordSet 要Upd 傳入的Recordset By Kin 2011/08/09
                        Dim blnUpdOld As Boolean
                        blnUpdOld = True
                        If (Not rsUpd Is Nothing) Then
                        
                            If (rsUpd.State = adStateOpen) Then
                                If rsUpd.RecordCount > 0 Then
                                    Dim aDate As String
                                    aDate = RightNow
                                    blnUpdOld = False
                                    rsUpd!CMBaudRate = varPara(12) & ""
                                    rsUpd!CMBaudNo = Get_BaudRate_Code(varPara(12) & "") & ""
                                    rsUpd!UpdEn = garyGi(1) & ""
                                    rsUpd!UpdTime = GetDTString(aDate) & ""
                                    rsUpd!CMOpenDate = aDate
                                    rsUpd.Update
                                    
                                    '#6066 測試不OK舊設備要UPD CMCLOSEDATE By Kin 2011/08/23
                                    '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                                    '判斷正常設備又要拿掉 For 銘誠 By Kin 2014/10/16
                                    strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                            " SET  UPDEN='" & garyGi(1) & "',UPDTIME='" & GetDTString(aDate) & "' ," & _
                                                            " CMCLOSEDATE = TO_DATE('" & Format(aDate, "yyyymmddhhmmss") & "','YYYYMMDD HH24:MI:SS') " & _
                                                            " WHERE SEQNO='" & strSeqNo & "'"
                                             
                                    gcnGi.Execute strUpdSQL
                                End If
                            End If
                        End If
                        
                        If blnUpdOld Then
                            If strPrcType = "" Then
                                If Val(rs307!CMBaudRateType & "") = 2 Then
                                    strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                            " SET  UPDEN='" & garyGi(1) & "',UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                            " WHERE SEQNO='" & strSeqNo & "'"
                                Else
                                    strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                            " SET CMBAUDRATE='" & varPara(12) & "'" & _
                                                            ",CMBAUDNO=" & Get_BaudRate_Code(varPara(12) & "") & "" & _
                                                            ",UPDEN='" & garyGi(1) & "',UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                            " WHERE SEQNO='" & strSeqNo & "'"
                                End If
                                gcnGi.Execute strUpdSQL
                            Else
                                rs!CMBaudRate = varPara(12) & ""
                                rs!CMBaudNo = Get_BaudRate_Code(varPara(12) & "") & ""
                                rs!UpdEn = garyGi(1) & ""
                                rs!UpdTime = GetDTString(RightNow) & ""
                                
                            End If
                        End If
                        
'                        戊、    依原升降速率之API (A1_BU)回寫資料庫並增加寫入如下欄位：(無預約日)
'                        1.  SO307 : CMBaudRateType速率變動識別=2值 (2=試用)。
'                        2.  SO004 : ProbationRateNo試用速率代碼= gilNewRate.GetCodeNo。
'                                           ProbationRateName試用速率= gilNewRate.GetDescription。
'                                           ProbationStopDate預計試用截止日= gdaProbationStopDate.GetValue。
'                                           ProbationStopFlag停止試用狀態=0。
'                        己、    依原升降速率之API (A1_BU)回寫資料庫並增加寫入如下欄位：(有預約日)
'                        1.  SO307 : CMBaudRateType速率變動識別=2值 (2=試用)。
                        
                        Select Case Val(rs307!CMBaudRateType & "")
                            Case 0 ' 派工升降
                            Case 1 ' 免費升降
                            Case 2 ' 試用
                            ' O2Date
                                        If UI_Mode And strBaudRate = "" Then
                                            strBaudRate = Screen.ActiveForm.gilNewRate.GetCodeNo & ""
                                        End If
                                        If strPrcType = "" Then
                                            If strBaudRate = "" Then strBaudRate = Get_BaudRate_Code(rs307!NewCMBaudRate & "")
                                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                                    " SET PROBATIONRATENO=" & strBaudRate & _
                                                                    ",PROBATIONRATENAME='" & rs307!NewCMBaudRate & "'" & _
                                                                    ",PROBATIONSTOPDATE=" & O2Date(rs307!ProbationStopDate & "", True) & _
                                                                    ",PROBATIONSTOPFLAG=0" & _
                                                                    " WHERE SEQNO='" & strSeqNo & "'"
                                            gcnGi.Execute strUpdSQL
                                        Else
                                            If strBaudRate = "" Then
                                                rs!ProbationRateNo = Get_BaudRate_Code(rs307!NewCMBaudRate & "")
                                            Else
                                                rs!ProbationRateNo = strBaudRate
                                            End If
                                            rs!ProbationRateName = rs307!NewCMBaudRate & ""
                                            rs!ProbationStopDate = rs307!ProbationStopDate & ""
                                            rs!ProbationStopFlag = 0
                                        End If
                            Case 3 ' 停止試用
                                        If strPrcType = "" Then
                                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                                    " SET PROBATIONSTOPDATE=" & O2Date(RightDate, True) & _
                                                                    ",PROBATIONSTOPFLAG=1" & _
                                                                    " WHERE SEQNO='" & strSeqNo & "'"
                                            gcnGi.Execute strUpdSQL
                                        Else
                                            rs!ProbationStopDate = RightDate
                                            rs!ProbationStopFlag = 1
                                        End If
                        End Select
                                                    
                    Case "A1_PIP" ' 申請動態IP
                    
                        If strPrcType = "" Then
                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                    " SET DYNIPCOUNT=DYNIPCOUNT+" & varPara(14) & _
                                                    ",UPDEN='" & garyGi(1) & "'" & _
                                                    ",UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                    " WHERE SEQNO='" & strSeqNo & "'"
                                                        
                            gcnGi.Execute strUpdSQL
                        Else
                            rs!DYNIPCOUNT = Val(rs!DYNIPCOUNT & "") + Val(varPara(14) & "")
                        End If
                        
                    Case "A1_MIP" ' 取消動態IP
                    
                        If strPrcType = "" Then
                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                    " SET DYNIPCOUNT=DYNIPCOUNT" & varPara(14) & _
                                                    ",UPDEN='" & garyGi(1) & "'" & _
                                                    ",UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                    " WHERE SEQNO='" & strSeqNo & "'"
                                                        
                            gcnGi.Execute strUpdSQL
                        Else
                            rs!DYNIPCOUNT = Val(rs!DYNIPCOUNT & "") - Abs(Val(varPara(14) & ""))
                        End If
                        
                    Case "A2" ' 申請固定IP
                        '指令執行成功需新增一筆到so004c (請按照rd原結構關聯對應資料新增)，
                        '且更新so004.的固定ip 值（舊的ip值＋新增的ip值）
                        varAry = Split(varPara(23), "#")
                        varAry2 = Split(varPara(26), "#")
                        For bytLoop = 0 To UBound(varAry)
                            strUpdSQL = "INSERT INTO " & GetOwner & "SO004C" & _
                                                    " (CUSTID,FACISNO,SEQNO,CPEMAC,IPADDRESS) VALUES" & _
                                                    " (" & varPara(6) & "," & _
                                                    " '" & strFaciSno & "'," & _
                                                    " '" & strSeqNo & "'," & _
                                                    " '" & Replace(varAry(bytLoop), ":", "") & "'," & _
                                                    " '" & varAry2(bytLoop) & "')"
                            gcnGi.Execute strUpdSQL
                        Next
                        If strPrcType = "" Then
                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                    " SET FIXIPCOUNT=FIXIPCOUNT+" & varPara(16) & _
                                                    ",UPDEN='" & garyGi(1) & "'" & _
                                                    ",UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                    " WHERE SEQNO='" & strSeqNo & "'"
                            
                            gcnGi.Execute strUpdSQL
                        Else
                            rs!FixIPCount = Val(rs!FixIPCount & "") + Val(varPara(16) & "")
                        End If
                        
                    Case "A3" ' 取消固定IP
                        '指令執行成功需，且更新指定取消的so004c ip 與cpe mac 資料該筆紀錄stopdate=系統時間
                        'so0042的固定ip 值(舊的ip值 - 新增的ip值)
                        varAry = Split(varPara(24), "#")
                        varAry2 = Split(varPara(27), "#")
                        For bytLoop = 0 To UBound(varAry)
                            strUpdSQL = "UPDATE " & GetOwner & "SO004C" & _
                                                    " SET STOPDATE=" & O2Date(RightNow) & _
                                                    " WHERE FACISNO='" & strFaciSno & "'" & _
                                                    " AND SEQNO='" & strSeqNo & "'" & _
                                                    " AND CPEMAC='" & Replace(varAry(bytLoop), ":", "") & "'" & _
                                                    " AND IPADDRESS='" & varAry2(bytLoop) & "'" & _
                                                    " AND CUSTID=" & varPara(6)
                            'strUpdSQL = "DELETE FROM " & GetOwner & "SO004C" & _
                                                    " WHERE FACISNO='" & strFaciSno & "'" & _
                                                    " AND SEQNO='" & strSeqNo & "'" & _
                                                    " AND CPEMAC='" & Replace(varAry(bytLoop), ":", "") & "'" & _
                                                    " AND IPADDRESS='" & varAry2(bytLoop) & "'" & _
                                                    " AND CUSTID=" & varPara(6)
                            gcnGi.Execute strUpdSQL
                        Next
                        If strPrcType = "" Then
                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                    " SET FIXIPCOUNT=FIXIPCOUNT" & varPara(16) & _
                                                    ",UPDEN='" & garyGi(1) & "'" & _
                                                    ",UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                    " WHERE SEQNO='" & strSeqNo & "'"
                            gcnGi.Execute strUpdSQL
                        Else
                            rs!FixIPCount = Val(rs!FixIPCount & "") - Abs(Val(varPara(16) & ""))
                        End If
                        
                    Case "A4", "E6" ' EMTA IP異動
                        ' 若選取catv 移機中是否依據移機新址選取時則網編尋找條件（按照so041取用nodeno或circuitno）
                        ' 查詢出可用的ip，ip 被選取時別的user 不可選取，新node 亦須寫入資料交換table，
                        
                        ' 指令若成功則更新so004的ipaddress 與ip 取用table的使用狀態
                        ' 此ip，指令若成功則更新so004的ipaddress 與ip 取用table的使用狀態，
                        
                        ' 增加SO049 IP 類別判斷屬於CM 的才可以選取使用。
                        ' (2006/06/21 指令成功沒有將so049 IP 設成已使用,so004 的ipaddress 也沒有做更新)
                        
                        If strPrcType = "" Then
                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                    " SET IPADDRESS='" & varPara(29) & "'" & _
                                                    " WHERE SEQNO='" & strSeqNo & "'" & _
                                                    " AND CUSTID=" & varPara(6) & _
                                                    " AND COMPCODE=" & varPara(5)
                            
                            gcnGi.Execute strUpdSQL
                        Else
                            rs!IPADDRESS = varPara(29) & ""
                        End If
                        
                        strUpdSQL = "UPDATE " & GetOwner & "SO048" & _
                                                " SET USEFLAG=1 WHERE IPADDRESS='" & varPara(29) & "'" & _
                                                " AND COMPCODE=" & varPara(5)
                                                
                        gcnGi.Execute strUpdSQL
                        
                        If Not Screen.ActiveForm Is Nothing Then Screen.ActiveForm.strIPaddr = ""
                    Case "A5", "A6" ' CPE IP 異動 / CPE MAC 異動
                        ' A5 : 指令成功將舊ip STOPDATE 填入當時系統時間,複製一樣的資料ip 不同新增到so004C
                        ' A6 : 指令成功將舊ip STOPDATE填入當時系統時間,複製一樣的資料cpe mac 不同新增到so004C。
                        
                        'If strPrcType = "" Then
                                    
                            strUpdSQL = "UPDATE " & GetOwner & "SO004C" & _
                                                    " SET STOPDATE=" & O2Date(RightNow) & _
                                                    " WHERE FACISNO='" & strFaciSno & "'" & _
                                                    " AND SEQNO='" & strSeqNo & "'" & _
                                                    " AND CUSTID=" & varPara(6)
                            
                            gcnGi.Execute strUpdSQL
                            
                            strUpdSQL = "INSERT INTO " & GetOwner & "SO004C" & _
                                                    " (CUSTID,FACISNO,SEQNO,CPEMAC,IPADDRESS) VALUES" & _
                                                    " (" & varPara(6) & "," & _
                                                    " '" & strFaciSno & "'," & _
                                                    " '" & strSeqNo & "'," & _
                                                    " '" & Replace(varPara(23), ":", "") & "'," & _
                                                    " '" & varPara(26) & "')"
                            
                            gcnGi.Execute strUpdSQL
                            
                        'End If
                                                
                    Case "A7", "A8", "A9", "A10" ' CM / EMTA 更換
                        '品名右側輸入新cm 的mac
                        '指令成功後新增一筆設備異動紀錄
                        '紀錄新舊設備品名與序號
                        '更新so004C.FaciSno與舊序號有關聯的內容
                        'so003 , so033關聯此設備的收費資料異動關聯值
                        If strPrcType <> "" Then
                             '#6813 增加判斷正常設備才要更新 By Kin 2014/06/18
                             '判斷正常設備又要拿掉 For 銘誠 By Kin 2014/10/16
                            rs!CMCloseDate = RightNow
                            If Not rsUpd Is Nothing And Not IsMissing(rsUpd) Then
                                If rsUpd.State > 0 Then
                                    rsUpd!CMOpenDate = RightNow
                                    rsUpd.Update
                                End If
                            End If
                        Else
                            strUpdSQL = "UPDATE " & GetOwner & "SO003" & _
                                                    " SET FACISNO='" & Replace(varPara(8), ":", "") & "'" & _
                                                    " WHERE FACISEQNO='" & strSeqNo & "'" & _
                                                    " AND CUSTID=" & varPara(6) & _
                                                    " AND COMPCODE=" & varPara(5)
                            
                            gcnGi.Execute strUpdSQL
                                                    
                            strUpdSQL = "UPDATE " & GetOwner & "SO004C" & _
                                                    " SET FACISNO='" & Replace(varPara(8), ":", "") & "'" & _
                                                    " WHERE SEQNO='" & strSeqNo & "'" & _
                                                    " AND CUSTID=" & varPara(6)
                            
                            gcnGi.Execute strUpdSQL
                                
    '                                                " ,CPEMAC='" & varAry(bytLoop) & "'" & _
    '                                                " ,IPADDRESS='" & varAry2(bytLoop) & "'" & _

                            ' Update SO004 's ModelCode , ModelName , FaciCode , FaciName
                            ' 指令傳送需要按照新的序號到SO306去取新的設備型號代碼再取得型號指令用代碼作傳送,
                            ' 指令成功後需要將SO004設備品名與型號與序號更新成新的
                            
                            With Screen
                                If Not .ActiveForm Is Nothing Then ' from console panel
                                    strFC = .ActiveForm.gilFaciCode.GetCodeNo & ""
                                    strFN = .ActiveForm.gilFaciCode.GetDescription & ""
                                End If
                            End With
                                
                            strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                    " SET FACISNO='" & Replace(varPara(8), ":", "") & "'" & _
                                                    ",FACICODE=" & strFC & _
                                                    ",FACINAME='" & strFN & "'" & _
                                                    ",MODELCODE=" & strMC & _
                                                    ",MODELNAME='" & strMN & "'" & _
                                                    " WHERE SEQNO='" & strSeqNo & "'" & _
                                                    " AND CUSTID=" & varPara(6) & _
                                                    " AND COMPCODE=" & varPara(5)
                            
                            gcnGi.Execute strUpdSQL
                            
                            If varPara(4) = "A8" Then
                                
                                strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                        " SET REFACISNO='" & Replace(varPara(8), ":", "") & "'" & _
                                                        " WHERE REFACISNO='" & strFaciSno & "'" & _
                                                        " AND CUSTID=" & varPara(6) & " AND COMPCODE=" & varPara(5) & _
                                                        " AND SERVICETYPE='P' AND PRDATE IS NULL" & _
                                                        " AND FACICODE IN (" & _
                                                        " SELECT CODENO FROM " & GetOwner & "CD022" & _
                                                        " WHERE REFNO=6)"
                                                    
                                gcnGi.Execute strUpdSQL
                                
                            End If
                        
    
                            If varPara(4) = "A9" Then
                                ' 參考號2舊CM /換新參考號5 EMTA A9 需要自動依據NODE 取新的IPADDRESS，
                                
                                If strPrcType = "" Then
                                    strUpdSQL = "UPDATE " & GetOwner & "SO004" & _
                                                            " SET IPADDRESS='" & varPara(29) & "'" & _
                                                            " WHERE SEQNO='" & strSeqNo & "'" & _
                                                            " AND CUSTID=" & varPara(6) & _
                                                            " AND COMPCODE=" & varPara(5)
                                                            
                                    gcnGi.Execute strUpdSQL
                                Else
                                    rs!IPADDRESS = varPara(29) & ""
                                End If
                                
                                ' 指令成功後需更新SO004.IPADDRESS 與將IP USEFLAG 改成已使用
                                
                                strUpdSQL = "UPDATE " & GetOwner & "SO048" & _
                                                        " SET USEFLAG=1 WHERE IPADDRESS='" & varPara(29) & "'" & _
                                                        " AND COMPCODE=" & varPara(5)
                                
                                gcnGi.Execute strUpdSQL
                                                        
                            End If
                            ' 2006/06/21 輸入新序號時需要檢核該序號在SO004 中是否有重複(沒有拆機日期的)及算重複不可作設定,
 
                        End If
                        
                    Case "Q1" ' 重設 CM Reset
'                        gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
                                                " SET CMOPENDATE=" & O2Date(RightNow) & _
                                                ",UPDEN='CABLESOFT',UPDTIME='" & GetDTString(RightNow) & "'" & _
                                                " WHERE FACISNO='" & strFaciSno & "'"
                End Select
            End If
        End If
        
        If CmdSeqNo <> Empty Then Exit Function
        CmdSeqNo = strCmdSno
        
'       網路編號 / NODE / VOIP
'       gcnGi.Execute "UPDATE " & GetOwner & "SO004" & _
                                    " SET IPADDRESS='" & str_IP_Address & "'" & _
                                    ",UPDEN='CABLESOFT'" & _
                                    ",UPDTIME='" & GetDTString(RightNow) & "'" & _
                                    " WHERE REFACISNO='" & Replace(str_CM_MAC, ":", "") & "'" & _
                                    " AND COMPCODE=" & strCompCode & _
                                    " AND CUSTID=" & str_Cust_ID & _
                                    " AND SERVICETYPE='I'" & _
                                    " AND PRDATE IS NULL"
        
'        ",FIXIPCOUNT=" & str_Fix_IP_Count & _

    Else
        Ins2SwapTbl = False
    End If
    
    ' 要寫 T
    ' W -> C -> S
    If blnTimeOut Then ' Console high level commad time out
    
        strUpdSQL = "UPDATE " & strCMowner & "SO307 SET COMMANDSTATUS='T'" & _
                                " WHERE CMDSEQNO='" & strCmdSno & "'"
                                
        objCmdCN.Execute strUpdSQL
        
    End If
    
    If Not Screen.ActiveForm Is Nothing Then ' from console panel
        With Screen.ActiveForm
                GetRS307 , strCmdSno ' Requery data
                .ShowResult rs307, .ggrQueryInfo ' Show Resutl
                .pbr1.Max = 1
                .pbr1.Value = 1
        End With
    End If
    
    ' Status 成功 , 一樣要 Show Msg ( RetCode , RetMsg )
'    If (UCase(strStatus) <> "W" And UCase(strStatus) <> "C") Or blnTimeOut Then
    strErrMsg = Empty
    Select Case UCase(strStatus)
        Case "S"
'            strErrMsg = "[ " & strCmdName & " ] 控制訊號已送出完成 !! "
            strErrMsg = "[ " & strCmdName & " ] 執行成功 ! "
        Case "W", "C"
            strErrMsg = "[ " & strCmdName & " ] 執行失敗 ! " & strCrLf & "作業逾時 ( Time out ) !"
        Case Else
            strErrMsg = "[ " & strCmdName & " ] 執行失敗 ! "
'            strErrMsg = "[ " & strCmdName & " ] 失敗 !! " & strCrlf & GetErrMsg(CStr(lngReturn))
'            If strRetMsg <> Empty Then strErrMsg = strErrMsg & strCrlf & strRetMsg
'        Case "E"
'            strErrMsg = "[ " & strCmdName & " ] 失敗 !! " & strCrlf & GetErrMsg(CStr(lngReturn))
'            If strRetMsg <> Empty Then strErrMsg = strErrMsg & strCrlf & strRetMsg
'        Case "GT" ' Gateway 與 IP Command time out !
'        Case "TR"
'            strErrMsg = "[ " & strCmdName & " ] 失敗 !! " & strCrlf & GetErrMsg(CStr(lngReturn))
'            If strRetMsg <> Empty Then strErrMsg = strErrMsg & strCrlf & strRetMsg
'        Case "TE"
'            strErrMsg = "[ " & strCmdName & " ] 失敗 !! " & strCrlf & GetErrMsg(CStr(lngReturn))
'            If strRetMsg <> Empty Then strErrMsg = strErrMsg & strCrlf & strRetMsg
    End Select
    
    If strPrcType <> "" Then rs.Update
    
    strErrMsg = strErrMsg & strCrLf & GetErrMsg(CStr(lngReturn))
    If strRetMsg <> Empty Then strErrMsg = strErrMsg & strCrLf & strRetMsg
    If UI_Mode Then MsgBox strErrMsg, vbInformation, gimsgPrompt
    AddRetMsg strErrMsg
'    End If
    
  Exit Function
ChkErr:
    ErrSub FrmName, "Ins2SwapTbl"
    If Len(Dir("C:\Debug.txt")) > 0 Then
        CreateObject("Scripting.FileSystemObject").CreateTextFile("C:\Debug.txt").Write _
            strTrcPos & vbCrLf & strUpdSQL & vbCrLf & strErrMsg
        Shell "Notepad C:\Debug.txt", vbNormalNoFocus
'        MsgBox strTrcPos & vbCrLf & strUpdSQL, , "Debug Mode"
    End If
End Function

' #4302
' (RA)_SO1171A_20081223_調整CLI控制台CP門號判斷.doc
' 2009/01/03 By Power Hammer
Public Function IsCP(ByRef rs04 As Object) As Boolean
  On Error GoTo ChkErr
    ' 門號匯入模式 : 0-台固、1-亞太、2-速博
    IsCP = GetRsValue("Select Count(*) From " & GetOwner & "CD022" & _
                        " Where CodeNo=" & rs04!FaciCode & _
                        " And CPImportMode In (0,1)") > 0
  Exit Function
ChkErr:
    ErrSub "mod_ControlPanel", "IsCP"
End Function

Private Function Get_BaudRate_Code(ByRef BaudRateDesc As String) As String
  On Error GoTo ChkErr
    Get_BaudRate_Code = GetRsValue("SELECT CODENO FROM " & GetOwner & "CD064 WHERE DESCRIPTION='" & BaudRateDesc & "'", gcnGi) & ""
  Exit Function
ChkErr:
    ErrSub "mod_ControlPanel", "Get_BaudRate_Code"
End Function

Public Function GetNewIP() As String
  On Error Resume Next
    GetNewIP = GetRsValue("SELECT IPADDRESS FROM " & GetOwner & "SO048" & _
                                            " WHERE CIRCUITNO = " & _
                                            "(SELECT " & IIf(int_EMTA_IP_Type = 0, "CIRCUITNO", "NODENO") & _
                                            " FROM " & GetOwner & "SO014" & _
                                            " WHERE ADDRNO = " & lngInstAddrNo & _
                                            " AND COMPCODE = " & garyGi(9) & ") " & _
                                            " AND COMPCODE=" & garyGi(9) & " AND USEFLAG=0") & ""
    If Err Then GetNewIP = ""
End Function

Public Sub ShowVtcFld(ctlGrd As Control)
  On Error GoTo ChkErr
    Dim objGrd As Object
    Set objGrd = CreateObject("csGridFieldView.clsGridFiledView")
    With objGrd
        Set .uSourceGrid = ctlGrd
        Set .uSourceRS = ctlGrd.Recordset
        Set .uGridFields = mFlds
        .uWidth1 = 3456
        .uWidth2 = 3456
        .ShowGridSortField 1
    End With
  Exit Sub
ChkErr:
    ErrSub FrmName, "ShowVtcFld"
End Sub

Public Function GetRS307(Optional strFlds As String = "", Optional strKey As String = "") As Boolean
  On Error GoTo ChkErr
    GetRS307 = GetRS(rs307, "SELECT " & IIf(strFlds = "", "*", strFlds) & " FROM " & strCMowner & "SO307" & _
                                                " WHERE " & IIf(strKey = "", "0=1", "CMDSEQNO='" & strKey & "'"), objCmdCN, 3, 3, 3)
  Exit Function
ChkErr:
    ErrSub FrmName, "GetRS307"
End Function

Private Function FmtMac(ByVal strMac As String) As String
  On Error GoTo ChkErr
    Dim bytLoop As Byte
    If strMac <> Empty Then
        For bytLoop = 1 To 12 Step 2
            FmtMac = FmtMac & ":" & Mid(strMac, bytLoop, 2)
        Next
        FmtMac = Mid(FmtMac, 2)
    End If
  Exit Function
ChkErr:
    ErrSub FrmName, "FmtMac"
End Function

Public Function Mask(strMac As String) As String
  On Error GoTo ChkErr
    Dim varAry, varElement
    Dim strItem() As String
    Dim bytLoop As Byte
    strMac = Replace(strMac, ":", "", 1)
    If InStr(1, strMac, "#") Then
        varAry = Split(strMac, "#")
        ReDim strItem(UBound(varAry)) As String
        For Each varElement In varAry
            strItem(bytLoop) = FmtMac(varElement)
            bytLoop = bytLoop + 1
        Next
        Mask = Join(strItem, "#")
    Else
        Mask = FmtMac(strMac)
    End If
  Exit Function
ChkErr:
    ErrSub FrmName, "Mask"
End Function

'EMTA 序號關聯到so004.servicetype='P' and prdate is null 的取出 p 服務的 so004.Gateway 取一筆即可
Private Function GetMediaGw(strFaciSno As String, lngCustId As Long) As String
  On Error Resume Next
    GetMediaGw = GetRsValue("SELECT GATEWAY FROM " & GetOwner & "SO004" & _
                                                " WHERE REFACISNO='" & strFaciSno & "'" & _
                                                " AND CUSTID=" & lngCustId & _
                                                " AND COMPCODE=" & garyGi(9) & _
                                                " AND SERVICETYPE='P'" & _
                                                " AND PRDATE IS NULL" & _
                                                " GROUP BY GATEWAY", gcnGi, "") & ""
    If Err Then
        Err.Clear
        GetMediaGw = ""
    End If
End Function

' CP 的電話門號（第一線＃第二線
' EMTA 序號關聯到REFACISNO and so004.servicetype='P' and prdate is null 的
' 取出 p 服務的so004.facisno 多筆時以#區隔"

' New 2006/09/01 ADD BY JIM
' 資料取值規則CP 的電話門號（第一線＃第二線）
' EMTA 序號關聯到REFACISNO and so004.servicetype='P' and prdate is null 的
' 取出 p 服務的so004.facisno 多筆時以#區隔
' 因為是拆分線所以去除要拆除的cp門號，所以要送給交換檔請取p服務
' 安裝中 (instdate is null and prdate is null ) 或
' 正常的門號(instdate is not null and prdate is null and prsno is not null)
' 才需要送給交換檔

Private Function GetCPno(strFaciSno As String, lngCustId As Long, Optional blnOnlyOne As Boolean = False) As String
  On Error Resume Next
    If blnOnlyOne Then
        GetCPno = GetRsValue("SELECT FACISNO FROM " & GetOwner & "SO004" & _
                                                    " WHERE REFACISNO='" & strFaciSno & "'" & _
                                                    " AND CUSTID=" & lngCustId & " AND COMPCODE=" & garyGi(9) & _
                                                    " AND SERVICETYPE='P'" & _
                                                    " AND (INSTDATE IS NULL AND PRDATE IS NULL) " & _
                                                    " OR (INSTDATE IS NOT NULL AND PRDATE IS NULL AND PRSNO IS NOT NULL)")
    
'        GetCPno = GetRsValue("SELECT FACISNO FROM " & GetOwner & "SO004" & _
                                                    " WHERE REFACISNO='" & strFaciSno & "'" & _
                                                    " AND CUSTID=" & lngCustId & " AND COMPCODE=" & garyGi(9) & _
                                                    " AND SERVICETYPE='P'" & _
                                                    " AND (PRDATE IS NULL OR PRSNO IS NULL)")
    Else
        GetCPno = gcnGi.Execute("SELECT FACISNO FROM " & GetOwner & "SO004" & _
                                                    " WHERE REFACISNO='" & strFaciSno & "'" & _
                                                    " AND CUSTID=" & lngCustId & _
                                                    " AND COMPCODE=" & garyGi(9) & _
                                                    " AND SERVICETYPE='P'" & _
                                                    " AND PRDATE IS NULL AND PRSNO IS NULL").GetString(adClipString, , "", "#", "") & ""
    End If
    If Err = 0 Then
        GetCPno = RPxx(GetCPno)
        GetCPno = Replace(GetCPno, "##", "#", 1)
        If Right(GetCPno, 1) = "#" Then GetCPno = Left(GetCPno, Len(GetCPno) - 1)
    Else
        Err.Clear
        GetCPno = ""
    End If
End Function

' 該cm facisno 關聯 so004c. FaciSno 且 stopflag=0 的 取出so004c.ipaddress
' ff:ff:ff:ff:ff#ee:ee:ee:ee:ee（有＃多筆時GATEWAY 需要作多次）
Private Function GetIPaddr(strSeqNo As String, lngCustId As Long) As String
  On Error Resume Next
'    GetIPaddr = gcnGi.Execute("SELECT IPADDRESS FROM " & GetOwner & "SO004C" & _
                                                " WHERE FACISNO='" & strFaciSno & "'" & _
                                                " AND STOPDATE IS NULL").GetString(2, , "", "#", "")
                                                
    GetIPaddr = gcnGi.Execute("SELECT IPADDRESS FROM " & GetOwner & "SO004C" & _
                                                " WHERE SEQNO='" & strSeqNo & "'" & _
                                                " AND CUSTID=" & lngCustId & _
                                                " AND STOPDATE IS NULL" & _
                                                " AND ROWNUM=1").GetString(2, , "", "#", "")
                                                
'    GetIPaddr = gcnGi.Execute("SELECT IPADDRESS FROM " & GetOwner & "SO004C" & _
                                                " WHERE FACISNO='" & strFaciSno & "'" & _
                                                " AND STOPFLAG=0 " & _
                                                " AND CUSTID=" & frmSO1623A.ginCustId.Text).GetString(2, , "", "#", "")
    If Err = 0 Then
        GetIPaddr = RPxx(GetIPaddr)
        GetIPaddr = Replace(GetIPaddr, "##", "#", 1)
        If Right(GetIPaddr, 1) = "#" Then GetIPaddr = Left(GetIPaddr, Len(GetIPaddr) - 1)
    Else
        Err.Clear
        GetIPaddr = ""
    End If
End Function

' 該cm facisno 關聯 so004c. FaciSno 且 stopflag=0 取出 so004c.CPEMAC
' 192.168.1.1#192.168.1.10（有＃多筆時GATEWAY 需要作多次）
'Private Function GetCPEmac(strFaciSno As String) As String

Private Function GetCPEmac(strSeqNo As String, lngCustId As Long) As String
  On Error Resume Next
'    GetCPEmac = gcnGi.Execute("SELECT CPEMAC FROM " & GetOwner & "SO004C" & _
                                                " WHERE FACISNO='" & strFaciSno & "'" & _
                                                " AND STOPDATE IS NULL").GetString(2, , "", "#", "")
                                                
    GetCPEmac = gcnGi.Execute("SELECT CPEMAC FROM " & GetOwner & "SO004C" & _
                                                " WHERE SEQNO='" & strSeqNo & "'" & _
                                                " AND CUSTID=" & lngCustId & _
                                                " AND STOPDATE IS NULL" & _
                                                " AND ROWNUM=1").GetString(2, , "", "#", "")
    If Err = 0 Then
        GetCPEmac = RPxx(GetCPEmac)
        GetCPEmac = Replace(GetCPEmac, "##", "#", 1)
        If Right(GetCPEmac, 1) = "#" Then GetCPEmac = Left(GetCPEmac, Len(GetCPEmac) - 1)
    Else
        Err.Clear
        GetCPEmac = ""
    End If
End Function

' so014(CircuitNo / NodeNo) 依據系統參數so041.EMTAIPTYPE判斷(0=網路編號 , 1=Node)
Public Function GetHFCNode(lngCustId As Long, intComp As Integer, strServiceType As String, _
                            Optional blnChkAddrNo As Boolean = False) As String
  On Error Resume Next
    Dim intEMTAType As Integer
    Dim lngAddrNo As Long
    ' 4313
    ' 2009/03/13 By Power Hammer
    ' 1. 若SO002.SERVICETYPE=’C’且SO002.WIPCODE3　=　11　時,則去找CATV未完工的移機工單中新裝機地址的NODE點。
    If Val(GetRsValue("Select Count(*) Cnt From " & GetOwner & "SO002" & _
                        " Where CompCode=" & intComp & _
                        " And CustID=" & lngCustId & _
                        " And ServiceType='C'" & _
                        " And WipCode3=11", gcnGi)) > 1 Then
        '   SELECT　NODENO　from　SO014　WHERE　ADDRNO　IN　(
        '   SELECT　NEWCHARGEADDRNO　FROM　SO009　WHERE　CUSTID=5
        '   AND　FINTIME　IS　NULL
        '   AND　RETURNCODE　IS　NULL
        '   AND　NEWCHARGEADDRNO　IS　NOT　NULL
        '   AND　SERVICETYPE='C'
        '   AND　PRCODE　IN　(SELECT　CODENO　FROM　CD007　WHERE　REFNO　IN　(3)))
        lngAddrNo = Val(GetRsValue("Select ReInstAddrNo From " & GetOwner & "SO009" & _
                                    " Where CustID=" & lngCustId & _
                                    " And CompCode=" & intComp & _
                                    " And FinTime Is Null" & _
                                    " And ReturnCode Is Null" & _
                                    " And ReInstAddrNo Is Not Null" & _
                                    " And ServiceType='C'" & _
                                    " And PRCode In (Select CodeNo From " & GetOwner & "CD007 Where RefNo = 3)", _
                                    gcnGi) & "")
        blnChkAddrNo = lngAddrNo = 0
    Else
        lngAddrNo = Val(GetRsValue("SELECT INSTADDRNO FROM " & GetOwner & "SO001" & _
                                    " WHERE CUSTID = " & lngCustId & " AND COMPCODE = " & intComp, gcnGi) & "")
    End If
    If lngAddrNo <> 0 Then
        intEMTAType = GetSystemParaItem("EMTAIPTYPE", CStr(intComp), strServiceType, "SO041", , True)
        GetHFCNode = GetRsValue("SELECT " & IIf(intEMTAType = 0, "CIRCUITNO", "NODENO") & _
                                                    " FROM " & GetOwner & "SO014" & _
                                                    " WHERE ADDRNO = " & lngAddrNo & _
                                                    " AND COMPCODE = " & intComp) & ""
        If Err Then GetHFCNode = ""
    End If
  Exit Function
ChkErr:
    ErrSub FrmName, "GetHFCNode"
End Function

' Debby 查到 CM 時 , MTAMAC 會無值
' 由so004.facisno 關聯 so306.hfcmac 取出 so306.emtamac
Public Function GetMTAMAC(strFaciSno As String) As String
  On Error Resume Next
    If strFaciSno <> Empty Then
        GetMTAMAC = GetRsValue("SELECT MTAMAC FROM " & GetOwner & "SO306" & _
                                                    " WHERE HFCMAC='" & strFaciSno & "'", _
                                                    gcnGi, "[EMTA資料檔] 中查無 HFCMAC 為 [ " & strFaciSno & " ] 的資料 !") & ""
        If Err Then GetMTAMAC = ""
    End If
End Function

' 該台cm so004.ModelCode 對應到cd043.codeno 取cd043.CMModelCode
Private Function GetModel(strModelCode As String) As String
  On Error GoTo ChkErr ' CD043.CM ModelCode  (SO004.ModelCode JOIN)
    If strModelCode <> Empty Then
        GetModel = GetRsValue("SELECT CMMODELCODE FROM " & GetOwner & "CD043" & _
                                                            " WHERE CODENO=" & strModelCode, _
                                                            gcnGi, "[設備型號代碼檔] 中查無 MODELCODE 為 [ " & strModelCode & " ] 的資料 !") & ""
    End If
  Exit Function
ChkErr:
    ErrSub FrmName, "GetModel"
End Function

Public Function GetModelName(strFaciSno As String) As String
  On Error GoTo ChkErr
    Dim rs43 As New ADODB.Recordset
    If GetRS(rs43, "SELECT CODENO,DESCRIPTION,CMMODELCODE FROM " & GetOwner & "CD043 WHERE CODENO= " & _
                            "(SELECT MODELCODE FROM " & GetOwner & "SO306" & _
                            " WHERE HFCMAC='" & Replace(strFaciSno, ":", "") & "')") Then
        If Not rs43.EOF Then
            strMC = rs43!CodeNo & ""
            strMN = rs43!Description & ""
            GetModelName = rs43!CMModelCode & ""
        End If
    End If
    On Error Resume Next
    rs43.Close
    Set rs43 = Nothing
'    GetModelName = GetRsValue("SELECT CMMODELCODE FROM " & GetOwner & "CD043 WHERE CODENO= " & _
                                                    "(SELECT MODELCODE FROM " & GetOwner & "SO306" & _
                                                    " WHERE HFCMAC='" & Replace(strFaciSno, ":", "") & "')", gcnGi) & ""
  Exit Function
ChkErr:
    ErrSub FrmName, "GetModelName"
End Function

Private Function GetErrMsg(strRetCode As String) As String
  On Error Resume Next
    GetErrMsg = GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD080 WHERE CODENO='" & strRetCode & "'", gcnGi, "") & ""
'    GetErrMsg = GetRsValue("SELECT DESCRIPTION || CHR(13) || CHR(10) || EDESCRIPTION" & _
                                            " FROM " & GetOwner & "CD080 WHERE CODENO='" & strRetCode & "'", gcnGi, "")
    If Err.Number <> 0 Then
        Err.Clear
'        GetErrMsg = "錯誤代碼 : " & strRetCode & strFailureCode
        GetErrMsg = strRetCode & " " & strFailureCode
    Else
'        GetErrMsg = "錯誤名稱 : " & GetErrMsg & " " & strFailureCode
        GetErrMsg = GetErrMsg & " " & strFailureCode
    End If
End Function

' F - Failure , N - New ( Not Processed) , S - Success
Public Function Pooling(Optional ByRef objPGB As Object = Nothing, Optional strFunc As String) As Boolean
  On Error GoTo ChkErr
    
    Dim lngTimer As Long
    Dim intTickCount As Integer
    
    intTickCount = 0
    If intTimeOut = 0 Then GetSysPara
    
    If Not objPGB Is Nothing Then objPGB.Max = intTimeOut
    
    blnTimeOut = False
    If lngTimer = 0 Then lngTimer = GetTickCount
    
    Do Until intTickCount = intTimeOut
        If GetTickCount - lngTimer > 1000 Then
            lngTimer = GetTickCount
            If Not objPGB Is Nothing Then objPGB.Value = intTickCount
            DoEvents
            If Screen.ActiveForm Is Nothing Then
                Pooling = mod_ControlPanel.Qry_Result(False)
            Else
                Pooling = CallByName(Screen.ActiveForm, strFunc, VbMethod, False)
            End If
'            If UCase(strStatus) = "W" Or UCase(strStatus) = "C" Then Pooling = False
            If UCase(strStatus) <> "W" And UCase(strStatus) <> "C" Then Pooling = True
            If Pooling Then Exit Do
            intTickCount = intTickCount + 1
            If UI_Mode Then
                DoEvents
            End If
        End If
        If UI_Mode Then
            DoEvents
        End If
    Loop
    
    'If Not objPGB Is Nothing Then objPGB.Value = intTimeOut
    
    If Not Pooling Then
        If Screen.ActiveForm Is Nothing Then
            Pooling = mod_ControlPanel.Qry_Result(True)
        Else
            Pooling = CallByName(Screen.ActiveForm, strFunc, VbMethod, True)
        End If
    End If
    
    blnTimeOut = (UCase(strStatus) = "W") Or (UCase(strStatus) = "C")
    
  Exit Function
ChkErr:
    ErrSub FrmName, "Pooling"
End Function

'W：客服寫入待處理
'C: GateWay 已接收
'S: GateWay 執行成功
'E: GateWay 執行有錯誤
'T：客服系統 TimeOut GateWay 需要rollback
'GT：GatewWay 與ip command time out
'TR: GateWay rollback  成功 , TE:GateWay rollback  失敗"
Public Function Qry_Result(blnShowMsg As Boolean) As Boolean
  On Error GoTo ChkErr
    Dim rsResult As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim lngRcdAft As Long
    With cmd
            Set .ActiveConnection = objCmdCN
            .CommandType = 1
            .CommandText = "SELECT COMMANDSTATUS,RETCODE,RETMSG,CLIFAILURECODE" & _
                                        " FROM " & strCMowner & "SO307 WHERE CMDSEQNO=?"
            Set rsResult = .Execute(lngRcdAft, strCmdSno)
    End With
'    Debug.Print lngRcdAft
    With rsResult
        If Not .EOF Then
            strStatus = UCase(!CommandStatus & "")
            lngReturn = Val(!RetCode & "")
            strRetMsg = !RetMsg & ""
            strFailureCode = !cliFailureCode & ""
'            Debug.Print Switch(strStatus = "W", "客服寫入待處理", strStatus = "C", "Gateway 已接收", _
                                            strStatus = "S", "Gateway 執行成功", strStatus = "E", "Gateway 執行有錯誤")
            Qry_Result = (strStatus = "S" Or strStatus = "E")
'            Qry_Result = (strStatus = "S" Or strStatus = "E" Or strStatus = "GT")
        End If
    End With
    On Error Resume Next
    CloseRecordset rsResult
    Set cmd = Nothing
  Exit Function
ChkErr:
    ErrSub FrmName, "Qry_Result"
End Function

Private Sub PgbComplete()
  On Error Resume Next
    With Screen.ActiveForm.pbr1
        .Max = 1
        .Value = 1
    End With
End Sub

Public Sub GetSysPara()
  On Error GoTo ChkErr
    Dim rsPara As New ADODB.Recordset
    If GetRS(rsPara, "SELECT CMTIMEOUT,CMOWNER FROM " & GetOwner & "SO041" & _
                                " WHERE COMPCODE=" & garyGi(9), gcnGi, 2, 0, 1) Then
        intTimeOut = Val(rsPara("CMTimeOut") & "")
        strCMowner = rsPara("CMOWNER") & "."
    End If
'    intTimeOut = GetRsValue("SELECT CMTIMEOUT FROM " & GetOwner & "SO041 WHERE COMPCODE=" & garyGi(9))
    If intTimeOut = 0 Then intTimeOut = 30
    If LinkOdie Then strCMowner = "odie."

  Exit Sub
ChkErr:
    ErrSub FrmName, "GetSysPara"
End Sub
Public Function Get_BaudRate_Desc(ByRef BaudRateCode As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    
    strQry = "SELECT DESCRIPTION FROM " & GetOwner & "CD064 WHERE CODENO=" & BaudRateCode
    Get_BaudRate_Desc = GetRsValue(strQry, gcnGi) & ""
  Exit Function
ChkErr:
    ErrSub FrmName, "Get_BaudRate_Desc"
End Function

'Public Function Get_CPE_Mac(bytIP As Byte) As String
'  On Error GoTo ChkErr
'    Dim bytFixIP As Byte
'    bytFixIP = bytIP '
'    Select Case bytFixIP
'                Case 0
'                        Get_CPE_Mac = ""
'                Case 1
'                        Get_CPE_Mac = Get_CPE
'                        If Get_CPE_Mac = Empty Then
'                            Get_CPE_Mac = "NULL"
'                        Else
'                            If InStr(1, Get_CPE_Mac, "#") Then Get_CPE_Mac = Split(Get_CPE_Mac, "#")(0)
'                        End If
'                Case Else
'                        Get_CPE_Mac = Get_CPE
'                        Get_CPE_Mac = GetElement(Get_CPE_Mac, bytFixIP)
'    End Select
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Get_CPE_Mac"
'End Function

Private Function GetElement(strCPE As String, bytFixIP As Byte) As String
  On Error GoTo ChkErr
    Dim varAry As Variant
    Dim intLoop As Byte
    If InStr(1, strCPE, "#") Then
        varAry = Split(strCPE, "#")
        For intLoop = 0 To Val(bytFixIP) - 1
            On Error Resume Next
            GetElement = GetElement & "#" & varAry(intLoop)
            If Err.Number <> 0 Then
                Err.Clear
                GetElement = GetElement & "#" & "NULL"
            End If
        Next
        If Len(GetElement) > 2 Then GetElement = Mid(GetElement, 2)
    Else
        For intLoop = 1 To bytFixIP
            GetElement = GetElement & "#" & "NULL"
        Next
        GetElement = Mid(GetElement, 2)
    End If
  Exit Function
ChkErr:
    ErrSub FrmName, "GetElement"
End Function

'Public Function Get_CPE() As String
'  On Error GoTo ChkErr
'    Get_CPE = frmSO1623A.GetPCMac
'    If Right(Get_CPE, 1) = "#" Then Get_CPE = Left(Get_CPE, Len(Get_CPE) - 1)
'    If Get_CPE = "#" Then Get_CPE = ""
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "Get_CPE"
'End Function

'Public Function Get_VOIP_Type(strModelCode As String) As String
'  On Error GoTo ChkErr
'    If strModelCode <> Empty Then
'        Get_VOIP_Type = GetRsValue("SELECT VOICETYPE FROM " & GetOwner & "CD043" & _
'                                                        " WHERE CODENO=" & strModelCode, _
'                                                        gcnGi, "CD043 [設備型號代碼檔] 中查無 MODELCODE 為 [ " & strModelCode & " ] 的資料 !") & ""
'    End If
'  Exit Function
'ChkErr:
'    ErrSub FormName, "Get_VOIP_Type"
'End Function

'Public Function Get_Media_Gateway(strFaciSno As String) As String
'  On Error Resume Next ' SO004.FaciSno=SO049.CPNumber & CompCode => Gateway
'    Get_Media_Gateway = GetRsValue("SELECT GATEWAY FROM " & GetOwner & "SO049" & _
'                                                            " WHERE CPNUMBER='" & strFaciSno & "' AND" & _
'                                                            " COMPCODE=" & garyGi(9), gcnGi, "") & ""
'End Function

'Private Function GetEMTAtype() As Byte
'  On Error GoTo ChkErr
'    GetEMTAtype = Val(GetRsValue("SELECT EMTAIPTYPE FROM " & GetOwner & "SO041" & _
'                                                        " WHERE COMPCODE=" & garygi(9), gcnGi, "") & "")
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "GetEMTAtype"
'End Function


'Public Function GetSvcPvd() As String
'  On Error GoTo ChkErr
'    GetSvcPvd = GetRsValue("SELECT EDESCRIPTION FROM " & GetOwner & "CD039" & _
'                                            " WHERE CODENO=" & garygi(9), gcnGi, "") & ""
'    If GetSvcPvd = Empty Then GetSvcPvd = garygi(9)
'  Exit Function
'ChkErr:
'    ErrSub FrmName, "GetSvcPvd"
'End Function

' According FaciSno 2 Get SO306.MTAMAC
' RELATE SO004.FaciSno To SO306.HFCMAC
'Public Function Get_EMTA_Mac(strFaciSno As String) As String
'  On Error GoTo ChkErr
'    If strFaciSno <> Empty Then
'        Get_EMTA_Mac = GetRsValue("SELECT MTAMAC FROM " & GetOwner & "SO306" & _
'                                                        " WHERE HFCMAC='" & strFaciSno & "'", _
'                                                        gcnGi, "SO306 [EMTA資料檔] 中查無 HFCMAC 為 [ " & strFaciSno & " ] 的資料 !") & ""
'        Get_EMTA_Mac = Mask(Get_EMTA_Mac)
'    End If
'  Exit Function
'ChkErr:
'    ErrSub FormName, "Get_EMTA_Mac"
'End Function
